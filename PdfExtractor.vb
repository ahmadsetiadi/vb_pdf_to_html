' =====================================================================
'  PdfExtractor.vb — Tahap 2: baca teks + kotak/garis dari PDF (PdfPig)
'  Output: List(Of PageModel), koordinat mm, origin kiri-atas.
' =====================================================================
Imports System.Globalization
Imports UglyToad.PdfPig
Imports UglyToad.PdfPig.Content
Imports UglyToad.PdfPig.Core
Imports UglyToad.PdfPig.Graphics.Colors

Public Class PdfExtractor

    Private Const PtToMm As Double = 25.4 / 72.0

    Public Function Extract(pdfPath As String, log As Action(Of String)) As List(Of PageModel)
        Dim pages As New List(Of PageModel)

        Using doc = PdfDocument.Open(pdfPath)
            For Each page In doc.GetPages()
                Dim pm As New PageModel With {
                    .Number = page.Number,
                    .WidthMm = page.Width * PtToMm,
                    .HeightMm = page.Height * PtToMm
                }

                Dim words = page.GetWords().ToList()
                If words.Count = 0 Then
                    Throw New InvalidOperationException(
                        $"Halaman {page.Number} tidak mengandung teks. PDF harus berbasis teks (bukan hasil scan).")
                End If

                pm.Runs = BuildRuns(words, page.Height)
                pm.Shapes = BuildShapes(page, page.Height)

                log($"Halaman {page.Number}: {pm.WidthMm:0.#}x{pm.HeightMm:0.#} mm, {words.Count} kata → {pm.Runs.Count} run, {pm.Shapes.Count} shape")
                pages.Add(pm)
            Next
        End Using

        Return pages
    End Function

    ' ------------------------------------------------------------------
    '  TEKS: kata → baris → run (style seragam, jarak antar kata wajar)
    ' ------------------------------------------------------------------
    Private Class WordInfo
        Public Word As Word
        Public Baseline, Left, Right, Size As Double
        Public FontName, Color As String
        Public Bold, Italic As Boolean
    End Class

    Private Function BuildRuns(words As List(Of Word), pageHeightPt As Double) As List(Of TextRun)
        Dim runs As New List(Of TextRun)

        ' info per kata (style diambil dari huruf pertama)
        Dim infos As New List(Of WordInfo)
        For Each w In words
            Dim first = w.Letters(0)
            infos.Add(New WordInfo With {
                .Word = w,
                .Baseline = first.StartBaseLine.Y,
                .Left = w.BoundingBox.Left,
                .Right = w.BoundingBox.Right,
                .Size = first.PointSize,
                .FontName = first.FontName,
                .Bold = first.Font.IsBold OrElse GlobalSettings.IsBoldFont(first.FontName),
                .Italic = first.Font.IsItalic OrElse GlobalSettings.IsItalicFont(first.FontName),
                .Color = ColorToHex(first.Color)
            })
        Next
        infos = infos.OrderByDescending(Function(i) i.Baseline).ThenBy(Function(i) i.Left).ToList()

        ' kelompokkan jadi baris berdasarkan baseline (PDF: Y besar = atas)
        Dim lines As New List(Of List(Of WordInfo))
        Dim tol = GlobalSettings.BaselineTolerancePt
        Dim current As List(Of WordInfo) = Nothing
        Dim currentBaseline As Double = Double.NaN

        For Each i In infos
            If current Is Nothing OrElse Math.Abs(i.Baseline - currentBaseline) > tol Then
                current = New List(Of WordInfo)
                lines.Add(current)
                currentBaseline = i.Baseline
            End If
            current.Add(i)
        Next

        ' tiap baris → pecah jadi run (style sama & jarak antar kata wajar)
        For Each line In lines
            Dim run As TextRun = Nothing
            Dim prevRight As Double = 0

            For Each i In line.OrderBy(Function(x) x.Left)
                Dim spaceW = GlobalSettings.SpaceWidthEm * i.Size
                Dim gap = i.Left - prevRight
                Dim sameStyle = run IsNot Nothing AndAlso
                                Math.Abs(run.SizePt - i.Size) < 0.05 AndAlso
                                run.IsBold = i.Bold AndAlso run.IsItalic = i.Italic AndAlso
                                run.ColorHex = i.Color
                Dim closeEnough = run IsNot Nothing AndAlso gap >= -0.5 AndAlso gap <= GlobalSettings.MaxWordGapEm * spaceW

                If Not (sameStyle AndAlso closeEnough) Then
                    run = New TextRun With {
                        .XMm = i.Left * PtToMm,
                        .BaselineMm = (pageHeightPt - i.Baseline) * PtToMm,
                        .SizePt = i.Size,
                        .PdfFontName = i.FontName,
                        .IsBold = i.Bold,
                        .IsItalic = i.Italic,
                        .ColorHex = i.Color
                    }
                    runs.Add(run)
                End If

                run.Words.Add(New WordItem With {
                    .Text = i.Word.Text,
                    .XMm = i.Left * PtToMm,
                    .WidthMm = (i.Right - i.Left) * PtToMm
                })
                run.WidthMm = i.Right * PtToMm - run.XMm
                prevRight = i.Right
            Next
        Next

        Return runs
    End Function

    ' ------------------------------------------------------------------
    '  SHAPES: rect fill / rect stroke / path bebas (→ SVG)
    ' ------------------------------------------------------------------
    Private Function BuildShapes(page As Page, pageHeightPt As Double) As List(Of Shape)
        Dim shapes As New List(Of Shape)
        Dim inv = CultureInfo.InvariantCulture

        For Each path In page.ExperimentalAccess.Paths
            If Not path.IsFilled AndAlso Not path.IsStroked Then Continue For

            Dim fillHex = If(path.IsFilled, ColorToHex(path.FillColor), Nothing)
            Dim strokeHex = If(path.IsStroked, ColorToHex(path.StrokeColor), Nothing)
            Dim lineW As Double = path.LineWidth

            Dim allRects = path.All(Function(sp) sp.IsDrawnAsRectangle)

            If allRects Then
                For Each sp In path
                    Dim bb = sp.GetBoundingRectangle()
                    If Not bb.HasValue Then Continue For
                    Dim r = bb.Value
                    Dim s As New Shape With {
                        .XMm = r.Left * PtToMm,
                        .YMm = (pageHeightPt - r.Top) * PtToMm,
                        .WidthMm = r.Width * PtToMm,
                        .HeightMm = r.Height * PtToMm,
                        .FillHex = fillHex,
                        .StrokeHex = strokeHex,
                        .StrokeWidthPt = lineW
                    }
                    s.Kind = If(path.IsStroked AndAlso Not path.IsFilled, ShapeKind.StrokeRect, ShapeKind.FillRect)
                    shapes.Add(s)
                Next
            Else
                ' bentuk bebas (mis. logo) → SVG path, koordinat mm absolut
                Dim sb As New Text.StringBuilder()
                Dim bbAll = path.GetBoundingRectangle()
                For Each sp In path
                    For Each cmd In sp.Commands
                        If TypeOf cmd Is PdfSubpath.Move Then
                            Dim m = DirectCast(cmd, PdfSubpath.Move)
                            sb.Append(String.Format(inv, "M{0:0.###} {1:0.###} ", m.Location.X * PtToMm, (pageHeightPt - m.Location.Y) * PtToMm))
                        ElseIf TypeOf cmd Is PdfSubpath.Line Then
                            Dim l = DirectCast(cmd, PdfSubpath.Line)
                            sb.Append(String.Format(inv, "L{0:0.###} {1:0.###} ", l.To.X * PtToMm, (pageHeightPt - l.To.Y) * PtToMm))
                        ElseIf TypeOf cmd Is PdfSubpath.BezierCurve Then
                            Dim b = DirectCast(cmd, PdfSubpath.BezierCurve)
                            sb.Append(String.Format(inv, "C{0:0.###} {1:0.###} {2:0.###} {3:0.###} {4:0.###} {5:0.###} ",
                                b.FirstControlPoint.X * PtToMm, (pageHeightPt - b.FirstControlPoint.Y) * PtToMm,
                                b.SecondControlPoint.X * PtToMm, (pageHeightPt - b.SecondControlPoint.Y) * PtToMm,
                                b.EndPoint.X * PtToMm, (pageHeightPt - b.EndPoint.Y) * PtToMm))
                        ElseIf TypeOf cmd Is PdfSubpath.Close Then
                            sb.Append("Z ")
                        End If
                    Next
                Next
                If sb.Length = 0 Then Continue For

                Dim s As New Shape With {
                    .Kind = ShapeKind.Path,
                    .FillHex = fillHex,
                    .StrokeHex = strokeHex,
                    .StrokeWidthPt = lineW,
                    .SvgPathMm = sb.ToString().Trim()
                }
                If bbAll.HasValue Then
                    s.XMm = bbAll.Value.Left * PtToMm
                    s.YMm = (pageHeightPt - bbAll.Value.Top) * PtToMm
                    s.WidthMm = bbAll.Value.Width * PtToMm
                    s.HeightMm = bbAll.Value.Height * PtToMm
                End If
                shapes.Add(s)
            End If
        Next

        Return shapes
    End Function

    ' ------------------------------------------------------------------
    Public Shared Function ColorToHex(c As IColor) As String
        If c Is Nothing Then Return "#000000"
        Dim rgb = c.ToRGBValues()
        Dim r = CInt(Math.Round(CDbl(rgb.r) * 255))
        Dim g = CInt(Math.Round(CDbl(rgb.g) * 255))
        Dim b = CInt(Math.Round(CDbl(rgb.b) * 255))
        Return $"#{r:X2}{g:X2}{b:X2}"
    End Function

End Class
