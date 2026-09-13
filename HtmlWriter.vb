' =====================================================================
'  HtmlWriter.vb — Tahap 4: tulis page.css, header.html, footer.html,
'  bodyN.html, PageN.html, AllPages.html
' =====================================================================
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Text

Public Class HtmlWriter

    Private ReadOnly inv As CultureInfo = CultureInfo.InvariantCulture
    Private ReadOnly Utf8NoBom As New UTF8Encoding(False)   ' tanpa BOM supaya fragment bisa digabung
    Private pageWidthMm As Double
    Private pageHeightMm As Double

    Public Function Write(outDir As String, pages As List(Of PageModel), split As SplitResult, log As Action(Of String)) As List(Of String)
        Dim written As New List(Of String)
        Directory.CreateDirectory(outDir)
        pageWidthMm = pages(0).WidthMm
        pageHeightMm = pages(0).HeightMm

        ' ---------- page.css ----------
        Dim cssPath = Path.Combine(outDir, GlobalSettings.CssFileName)
        File.WriteAllText(cssPath, GlobalSettings.BuildPageCss(), Utf8NoBom)
        written.Add(cssPath)

        ' ---------- header.html / footer.html ----------
        Dim headerHtml = RenderRegion(split.Header, "hdr")
        Dim footerHtml = RenderRegion(split.Footer, "ftr")
        Dim headerPath = Path.Combine(outDir, GlobalSettings.HeaderFileName)
        Dim footerPath = Path.Combine(outDir, GlobalSettings.FooterFileName)
        File.WriteAllText(headerPath, headerHtml, Utf8NoBom) : written.Add(headerPath)
        File.WriteAllText(footerPath, footerHtml, Utf8NoBom) : written.Add(footerPath)

        ' ---------- bodyN.html + PageN.html ----------
        Dim all As New StringBuilder()
        all.AppendLine(HtmlHead("All Pages"))
        For i = 0 To pages.Count - 1
            Dim n = i + 1
            Dim bodyHtml = RenderRegion(split.Bodies(i), "bdy")
            Dim bodyPath = Path.Combine(outDir, String.Format(GlobalSettings.BodyFilePattern, n))
            File.WriteAllText(bodyPath, bodyHtml, Utf8NoBom) : written.Add(bodyPath)

            Dim pageNo = If(split.PageNumberTexts(i), "")
            Dim pageDiv = "<div class=""page"">" & vbLf &
                          headerHtml & vbLf & bodyHtml & vbLf &
                          footerHtml.Replace("{{PAGE}}", WebUtility.HtmlEncode(pageNo)) & vbLf &
                          "</div>"

            Dim pagePath = Path.Combine(outDir, $"Page{n}.html")
            File.WriteAllText(pagePath, HtmlHead($"Page {n}") & vbLf & pageDiv & vbLf & "</body></html>", Utf8NoBom)
            written.Add(pagePath)
            all.AppendLine(pageDiv)
        Next
        all.AppendLine("</body></html>")
        Dim allPath = Path.Combine(outDir, "AllPages.html")
        File.WriteAllText(allPath, all.ToString(), Utf8NoBom) : written.Add(allPath)

        Return written
    End Function

    Private Function HtmlHead(title As String) As String
        Return "<!DOCTYPE html>" & vbLf & "<html><head><meta charset=""utf-8""><title>" & WebUtility.HtmlEncode(title) &
               "</title><link rel=""stylesheet"" href=""" & GlobalSettings.CssFileName & """></head><body>"
    End Function

    ' ------------------------------------------------------------------
    '  Region → <div class="hdr|bdy|ftr" style="top:..;height:..">…</div>
    ' ------------------------------------------------------------------
    Private Function RenderRegion(r As Region, cssClass As String) As String
        Dim sb As New StringBuilder()
        sb.AppendLine(F("<div class=""{0}"" style=""top:{1}mm;height:{2}mm"">", cssClass, R2(r.TopMm), R2(r.HeightMm)))

        ' z-order: fill → stroke/path → teks
        For Each s In r.Shapes.Where(Function(x) x.Kind = ShapeKind.FillRect)
            sb.AppendLine("  " & RenderShape(s, r))
        Next
        For Each s In r.Shapes.Where(Function(x) x.Kind <> ShapeKind.FillRect)
            sb.AppendLine("  " & RenderShape(s, r))
        Next
        For Each run In r.Runs
            sb.AppendLine("  " & RenderRun(run, r))
        Next

        sb.Append("</div>")
        Return sb.ToString()
    End Function

    ' ------------------------------------------------------------------
    Private Function RenderShape(s As Shape, r As Region) As String
        Select Case s.Kind
            Case ShapeKind.StrokeRect
                ' kotak border → setting BorderColor / BorderWidthPt
                Return F("<div class=""r box"" style=""left:{0}mm;top:{1}mm;width:{2}mm;height:{3}mm""></div>",
                         R2(s.XMm), R2(s.YMm), R2(s.WidthMm), R2(s.HeightMm))

            Case ShapeKind.FillRect
                Dim thinPt = Math.Min(s.WidthMm, s.HeightMm) * 72 / 25.4
                If thinPt <= GlobalSettings.ThinLineMaxPt Then
                    ' garis tipis (grid tabel) → warna border
                    Return F("<div class=""r"" style=""left:{0}mm;top:{1}mm;width:{2}mm;height:{3}mm;background:{4}""></div>",
                             R2(s.XMm), R2(s.YMm), R2(s.WidthMm), R2(s.HeightMm), GlobalSettings.BorderColor)
                End If
                If IsTitleBar(s) Then
                    Return F("<div class=""r bar"" style=""left:{0}mm;top:{1}mm;width:{2}mm;height:{3}mm""></div>",
                             R2(s.XMm), R2(s.YMm), R2(s.WidthMm), R2(s.HeightMm))
                End If
                Return F("<div class=""r"" style=""left:{0}mm;top:{1}mm;width:{2}mm;height:{3}mm;background:{4}""></div>",
                         R2(s.XMm), R2(s.YMm), R2(s.WidthMm), R2(s.HeightMm), s.FillHex)

            Case Else ' Path → SVG selebar halaman, digeser -TopMm supaya koordinat absolut tetap benar
                Dim attrs As New StringBuilder()
                attrs.Append(If(s.FillHex IsNot Nothing, F("fill=""{0}"" ", s.FillHex), "fill=""none"" "))
                If s.StrokeHex IsNot Nothing Then
                    attrs.Append(F("stroke=""{0}"" stroke-width=""{1}"" ", s.StrokeHex, R2(s.StrokeWidthPt * 25.4 / 72)))
                End If
                Return F("<svg class=""r"" style=""left:0;top:{0}mm;width:{1}mm;height:{2}mm;overflow:visible"" viewBox=""0 0 {1} {2}""><path {3}d=""{4}""/></svg>",
                         R2(-r.TopMm), R2(pageWidthMm), R2(pageHeightMm), attrs.ToString(), s.SvgPathMm)
        End Select
    End Function

    Private Function IsTitleBar(s As Shape) As Boolean
        If s.FillHex Is Nothing Then Return False
        Return s.WidthMm >= GlobalSettings.TitleBarMinWidthMm AndAlso
               s.HeightMm >= 4 AndAlso s.HeightMm <= 12 AndAlso
               Luminance(s.FillHex) < 0.25
    End Function

    ' ------------------------------------------------------------------
    Private Function RenderRun(run As TextRun, r As Region) As String
        Dim size = If(GlobalSettings.UsePdfFontSize, run.SizePt, GlobalSettings.FontSizePt)
        Dim style As New StringBuilder()
        style.Append(GlobalSettings.TextStyle(run.IsBold, run.IsItalic, size))

        ' warna: teks gelap → FontColor setting; teks berwarna (merah/putih) → warna asli PDF
        If GlobalSettings.UsePdfTextColor OrElse Luminance(run.ColorHex) >= 0.3 OrElse IsColorful(run.ColorHex) Then
            style.Append("color:" & run.ColorHex & ";")
        End If

        ' word-spacing untuk baris justify: selisih lebar aktual vs natural dibagi jumlah spasi
        Dim n = run.Words.Count
        If n > 1 Then
            Dim natural = run.Words.Sum(Function(w) w.WidthMm) + (n - 1) * GlobalSettings.SpaceWidthEm * size * 25.4 / 72
            Dim extraMm = (run.WidthMm - natural) / (n - 1)
            Dim extraPt = extraMm * 72 / 25.4
            If Math.Abs(extraPt) > 0.05 Then style.Append(F("word-spacing:{0}pt;", R2(extraPt)))
        End If

        Dim ratio = If(run.IsItalic, GlobalSettings.BaselineRatioItalic, GlobalSettings.BaselineRatioNormal)
        Dim topMm = run.BaselineMm - ratio * size * 25.4 / 72

        Return F("<div class=""t"" style=""left:{0}mm;top:{1}mm;{2}"">{3}</div>",
                 R2(run.XMm), R2(topMm), style.ToString(), WebUtility.HtmlEncode(run.Text))
    End Function

    ' ------------------------------------------------------------------
    Private Function F(fmt As String, ParamArray args() As Object) As String
        Return String.Format(inv, fmt, args)
    End Function

    Private Function R2(v As Double) As String
        Return Math.Round(v, 2).ToString("0.##", inv)
    End Function

    Private Shared Function Luminance(hex As String) As Double
        If hex Is Nothing OrElse hex.Length < 7 Then Return 0
        Dim r = Convert.ToInt32(hex.Substring(1, 2), 16) / 255.0
        Dim g = Convert.ToInt32(hex.Substring(3, 2), 16) / 255.0
        Dim b = Convert.ToInt32(hex.Substring(5, 2), 16) / 255.0
        Return 0.2126 * r + 0.7152 * g + 0.0722 * b
    End Function

    Private Shared Function IsColorful(hex As String) As Boolean
        If hex Is Nothing OrElse hex.Length < 7 Then Return False
        Dim r = Convert.ToInt32(hex.Substring(1, 2), 16)
        Dim g = Convert.ToInt32(hex.Substring(3, 2), 16)
        Dim b = Convert.ToInt32(hex.Substring(5, 2), 16)
        Return Math.Max(r, Math.Max(g, b)) - Math.Min(r, Math.Min(g, b)) > 40
    End Function

End Class
