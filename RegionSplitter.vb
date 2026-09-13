' =====================================================================
'  RegionSplitter.vb — Tahap 3: bagi halaman jadi header / body / footer
'  Header & footer dideteksi dari kotak border (stroke rect) yang lebarnya
'  >= BoxMinWidthRatio x lebar halaman: paling atas = header, paling bawah = footer.
' =====================================================================
Imports System.Text.RegularExpressions

Public Class SplitResult
    Public Header As Region
    Public Footer As Region
    Public Bodies As New List(Of Region)          ' 1 per halaman
    Public PageNumberTexts As New List(Of String) ' teks nomor halaman tiap halaman (untuk {{PAGE}})
    Public HeaderBottomMm As Double
    Public FooterTopMm As Double
End Class

Public Class RegionSplitter

    Public Function Split(pages As List(Of PageModel), log As Action(Of String)) As SplitResult
        Dim res As New SplitResult
        Dim p1 = pages(0)

        ' ---------- deteksi batas header / footer dari halaman 1 ----------
        Dim wideBoxes = p1.Shapes.
            Where(Function(s) s.Kind = ShapeKind.StrokeRect AndAlso
                              s.WidthMm >= GlobalSettings.BoxMinWidthRatio * p1.WidthMm).
            OrderBy(Function(s) s.YMm).ToList()

        If wideBoxes.Count >= 2 Then
            Dim hb = wideBoxes.First()
            Dim fb = wideBoxes.Last()
            res.HeaderBottomMm = hb.YMm + hb.HeightMm + PtMm(hb.StrokeWidthPt)
            res.FooterTopMm = fb.YMm - PtMm(fb.StrokeWidthPt)
            log($"Kotak header terdeteksi: bawah = {res.HeaderBottomMm:0.##} mm; kotak footer: atas = {res.FooterTopMm:0.##} mm")
        Else
            res.HeaderBottomMm = GlobalSettings.HeaderHeightMm
            res.FooterTopMm = p1.HeightMm - GlobalSettings.FooterHeightMm
            log($"Kotak header/footer tidak terdeteksi → pakai fallback {GlobalSettings.HeaderHeightMm}/{GlobalSettings.FooterHeightMm} mm")
        End If

        ' ---------- header & footer dari halaman 1 ----------
        res.Header = Cut(p1, "header", 0, res.HeaderBottomMm)
        res.Footer = Cut(p1, "footer", res.FooterTopMm, p1.HeightMm)

        ' ---------- body tiap halaman ----------
        For Each p In pages
            res.Bodies.Add(Cut(p, $"body{p.Number}", res.HeaderBottomMm, res.FooterTopMm))
        Next

        ' ---------- teks isi yang meluber ke pita footer (template kepanjangan) → kembalikan ke body ----------
        If pages.Count >= 2 Then RescueOverflow(pages, res, log)

        ' ---------- nomor halaman → placeholder {{PAGE}} ----------
        DetectPageNumber(pages, res, log)

        Return res
    End Function

    ' Potong item halaman yang Y-tengahnya ada di [topMm, bottomMm); koordinat dibuat relatif.
    Private Function Cut(p As PageModel, name As String, topMm As Double, bottomMm As Double) As Region
        Dim r As New Region With {.Name = name, .TopMm = topMm, .HeightMm = bottomMm - topMm}

        For Each run In p.Runs
            Dim cy = run.BaselineMm - run.SizePt * 25.4 / 72 * 0.35   ' kira-kira tengah x-height
            If cy >= topMm AndAlso cy < bottomMm Then
                r.Runs.Add(Clone(run, topMm))
            End If
        Next

        For Each s In p.Shapes
            Dim cy = s.YMm + s.HeightMm / 2
            If cy >= topMm AndAlso cy < bottomMm Then
                r.Shapes.Add(Clone(s, topMm))
            End If
        Next

        Return r
    End Function

    ' Footer = teks yang berulang di posisi yang sama pada halaman lain. Run di pita footer yang
    ' tidak punya padanan di halaman mana pun (teks & posisi sama; angka boleh beda = nomor halaman)
    ' adalah isi body yang meluber ke bawah → dipindah ke body halaman itu (posisi y tetap, di bawah
    ' isi lain) supaya paragraf / <<endif>> tidak hilang dan mengalir ke halaman berikut saat dipecah.
    Private Sub RescueOverflow(pages As List(Of PageModel), res As SplitResult, log As Action(Of String))
        Dim footers = pages.Select(Function(p) Cut(p, "f", res.FooterTopMm, p.HeightMm)).ToList()
        ' teks per baris (baseline) tiap halaman, spasi dibuang — run boleh terpecah beda antar halaman ("Pemasar :" vs "Pemasar", ":")
        Dim lines = footers.Select(Function(f) f.Runs.GroupBy(Function(r) Math.Round(r.BaselineMm * 2) / 2).
                                       Select(Function(g) Tuple.Create(g.Key, String.Concat(g.OrderBy(Function(r) r.XMm).Select(Function(r) Squash(r.Text))))).ToList()).ToList()
        For pi = 0 To pages.Count - 1
            Dim moved As New List(Of String)
            For Each r In footers(pi).Runs
                Dim rr = r
                Dim key = Squash(rr.Text)
                Dim isNum = Regex.IsMatch(key, "^\d+$")
                Dim repeated = False
                For pj = 0 To pages.Count - 1
                    If pj = pi Then Continue For
                    If lines(pj).Any(Function(t) Math.Abs(t.Item1 - rr.BaselineMm) < 0.75 AndAlso (isNum OrElse key = "" OrElse t.Item2.Contains(key))) Then repeated = True : Exit For
                Next
                If repeated Then Continue For
                ' → body: koordinat relatif ke atas body
                Dim c = Clone(rr, 0)
                c.BaselineMm = rr.BaselineMm + res.FooterTopMm - res.HeaderBottomMm
                res.Bodies(pi).Runs.Add(c)
                moved.Add(rr.Text)
                If pi = 0 Then res.Footer.Runs.RemoveAll(Function(x) Math.Abs(x.BaselineMm - rr.BaselineMm) < 0.05 AndAlso Math.Abs(x.XMm - rr.XMm) < 0.05)
            Next
            If moved.Count > 0 Then log($"Halaman {pages(pi).Number}: {moved.Count} baris di area footer bukan bagian footer → ikut body: " &
                                        String.Join(" | ", moved.Select(Function(t) If(t.Length > 40, t.Substring(0, 40) & "…", t))))
        Next
    End Sub

    Private Shared Function Squash(t As String) As String
        Return Regex.Replace(t, "\s+", "")
    End Function

    ' Cari run di footer yang teksnya beda antar halaman (mis. "6" vs "9") → {{PAGE}}
    Private Sub DetectPageNumber(pages As List(Of PageModel), res As SplitResult, log As Action(Of String))
        Dim footers = pages.Select(Function(p) Cut(p, "f", res.FooterTopMm, p.HeightMm)).ToList()

        For Each p In pages
            res.PageNumberTexts.Add(Nothing)
        Next

        For idx = 0 To res.Footer.Runs.Count - 1
            Dim r1 = res.Footer.Runs(idx)
            Dim differs = False
            For pi = 1 To footers.Count - 1
                Dim other = footers(pi).Runs.FirstOrDefault(
                    Function(x) Math.Abs(x.BaselineMm - r1.BaselineMm) < 0.5 AndAlso Math.Abs(x.XMm - r1.XMm) < 1.5)
                If other Is Nothing OrElse other.Text <> r1.Text Then differs = True
            Next

            If differs AndAlso Regex.IsMatch(r1.Text, "^\d+$") Then
                For pi = 0 To footers.Count - 1
                    Dim other = footers(pi).Runs.FirstOrDefault(
                        Function(x) Math.Abs(x.BaselineMm - r1.BaselineMm) < 0.5 AndAlso Math.Abs(x.XMm - r1.XMm) < 1.5)
                    res.PageNumberTexts(pi) = If(other?.Text, "")
                Next
                r1.Words.Clear()
                r1.Words.Add(New WordItem With {.Text = "{{PAGE}}", .XMm = r1.XMm, .WidthMm = r1.WidthMm})
                log($"Nomor halaman terdeteksi di footer → {{{{PAGE}}}} (" & String.Join(", ", res.PageNumberTexts) & ")")
                Exit For
            End If
        Next

        ' kalau hanya 1 halaman / tidak ada yang beda → tidak ada placeholder
    End Sub

    ' ---------- helpers ----------
    Private Shared Function PtMm(pt As Double) As Double
        Return pt * 25.4 / 72
    End Function

    Private Shared Function Clone(run As TextRun, topMm As Double) As TextRun
        Dim c As New TextRun With {
            .XMm = run.XMm, .BaselineMm = run.BaselineMm - topMm, .WidthMm = run.WidthMm,
            .SizePt = run.SizePt, .PdfFontName = run.PdfFontName,
            .IsBold = run.IsBold, .IsItalic = run.IsItalic, .ColorHex = run.ColorHex
        }
        For Each w In run.Words
            c.Words.Add(New WordItem With {.Text = w.Text, .XMm = w.XMm, .WidthMm = w.WidthMm})
        Next
        Return c
    End Function

    Private Shared Function Clone(s As Shape, topMm As Double) As Shape
        Dim c As New Shape With {
            .Kind = s.Kind, .XMm = s.XMm, .YMm = s.YMm - topMm,
            .WidthMm = s.WidthMm, .HeightMm = s.HeightMm,
            .FillHex = s.FillHex, .StrokeHex = s.StrokeHex, .StrokeWidthPt = s.StrokeWidthPt,
            .SvgPathMm = s.SvgPathMm
        }
        Return c
    End Function

End Class
