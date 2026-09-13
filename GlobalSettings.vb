' =====================================================================
'  GlobalSettings.vb
'  Variable global untuk generator PDF -> HTML (header/footer/bodyN)
'  Semua nilai dipakai oleh HtmlWriter saat menulis CSS & inline style.
' =====================================================================
Imports System.Globalization

Public Module GlobalSettings

    ' ---------- Border (kotak header/footer/kolom, garis tabel) ----------
    Public BorderColor As String = "#bdb8b8"          ' warna border
    Public BorderWidthPt As Double = 0.3           ' tebal border (pt)

    ' ---------- Teks ----------
    Public FontColor As String = "#1A1A1A"         ' hitam agak gelap
    Public FontSizePt As Double = 8                ' ukuran font default (pt)

    ' ---------- Bar judul (PENGECUALIAN / BIAYA / PILIHAN DANA INVESTASI) ----------
    Public TitleBackgroundColor As String = "#00A758"   ' green (hijau Manulife)
    Public TitleFontColor As String = "#FFFFFF"

    ' ---------- Font name per style (rule dari PDF) ----------
    Public FontNameNormal As String = "Arial"
    Public FontNameItalic As String = "Monotype Corsiva"  ' teks italic di PDF -> font ini, style tetap normal
    Public FontNameBold As String = "Arial"
    Public FontFallback As String = "Arial, Helvetica, sans-serif"   ' dipakai jika font utama tidak terpasang

    ' ---------- Halaman ----------
    Public PageWidthMm As Double = 210
    Public PageHeightMm As Double = 297

    ' ---------- Output ----------
    Public OutputMode As String = "semantic"       ' "semantic" = struktur seperti riplay.html (cols/ol/table); "positional" = top/left per teks
    Public OutputFolder As String = "Output"
    Public HeaderFileName As String = "header.html"
    Public FooterFileName As String = "footer.html"
    Public BodyFilePattern As String = "body{0}.html"  ' body1.html, body2.html, ...
    Public CssFileName As String = "page.css"

    ' ---------- Fase 2: jarak saat menyatukan halaman (Generate PDF) ----------
    Public HeaderBodyGapMm As Double = 2       ' jarak border bawah kotak header → teks pertama body (margin-top blok pertama tiap halaman dibuang)
    Public BodyFooterGapMm As Double = 2       ' jarak minimum teks terakhir body → border atas kotak footer
    Public BreakBetweenSourcePages As Boolean = False  ' Generate Riplay: False = isi mengalir antar halaman sumber (halaman baru hanya di <<Page Break>>); True = tiap PageN.html mulai di halaman baru

    ' ---------- Fallback pemisah region (dipakai jika kotak header/footer tidak terdeteksi) ----------
    Public HeaderHeightMm As Double = 22
    Public FooterHeightMm As Double = 35

    ' ---------- Toleransi ekstraksi ----------
    Public BaselineTolerancePt As Double = 1.0     ' kata dianggap 1 baris jika selisih baseline <= ini
    Public BoxMinWidthRatio As Double = 0.9        ' rect dianggap kotak header/footer jika lebar >= 90% halaman
    Public SpaceWidthEm As Double = 0.278          ' lebar spasi Arial (em) – untuk hitung word-spacing justify
    Public MaxWordGapEm As Double = 3.0            ' jarak antar kata > 3x spasi → dianggap run terpisah
    Public ThinLineMaxPt As Double = 1.2           ' rect fill setipis ini dianggap garis (tabel) → pakai BorderColor
    Public UsePdfFontSize As Boolean = True        ' True = ukuran font ikut PDF; False = semua pakai FontSizePt
    Public UsePdfTextColor As Boolean = False      ' False = teks gelap dipaksa FontColor; teks berwarna (merah/putih) tetap asli
    Public TitleBarMinWidthMm As Double = 60       ' rect fill gelap selebar ini + tinggi 4–12mm → bar judul (green)
    Public DebugBands As Boolean = False            ' log rinci deteksi blok 2 kolom (SemanticHtmlWriter.DetectBands)

    ' Posisi baseline relatif dari atas div (line-height:1): Arial ascent 0.905, descent 0.212
    ' → (1 - 1.117)/2 + 0.905 = 0.8465. Corsiva perkiraan.
    Public BaselineRatioNormal As Double = 0.8465
    Public BaselineRatioItalic As Double = 0.85

    ' =====================================================================
    '  RULE FONT
    '  PDF normal  -> fontname = Arial,            fontstyle = normal
    '  PDF italic  -> fontname = Monotype Corsiva, fontstyle = normal (tetap normal)
    '  PDF bold    -> fontname = Arial,            fontstyle = bold
    ' =====================================================================
    Public Structure FontSpec
        Public FontName As String
        Public FontStyle As String    ' "normal" | "italic"
        Public FontWeight As String   ' "normal" | "bold"
    End Structure

    Public Function ResolveFont(isBold As Boolean, isItalic As Boolean) As FontSpec
        Dim f As FontSpec
        If isItalic Then
            f.FontName = FontNameItalic
            f.FontStyle = "normal"                 ' tetap normal walau di PDF italic
            f.FontWeight = If(isBold, "bold", "normal")
        ElseIf isBold Then
            f.FontName = FontNameBold
            f.FontStyle = "normal"
            f.FontWeight = "bold"
        Else
            f.FontName = FontNameNormal
            f.FontStyle = "normal"
            f.FontWeight = "normal"
        End If
        Return f
    End Function

    ' Deteksi bold/italic dari nama font di PDF (mis. "Arial-BoldMT", "Arial-ItalicMT", "ABCDEF+Arial,BoldItalic")
    Public Function IsBoldFont(pdfFontName As String) As Boolean
        Return pdfFontName.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Public Function IsItalicFont(pdfFontName As String) As Boolean
        Return pdfFontName.IndexOf("italic", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               pdfFontName.IndexOf("oblique", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    ' =====================================================================
    '  HELPER: inline style teks siap pakai
    ' =====================================================================
    Public Function TextStyle(isBold As Boolean, isItalic As Boolean, Optional sizePt As Double = -1) As String
        Dim f = ResolveFont(isBold, isItalic)
        Dim size = If(sizePt > 0, sizePt, FontSizePt)
        Return String.Format(CultureInfo.InvariantCulture,
            "font-family:'{0}',{5};font-style:{1};font-weight:{2};font-size:{3}pt;color:{4};",
            f.FontName, f.FontStyle, f.FontWeight, size, FontColor, FontFallback)
    End Function

    Public Function BorderStyle() As String
        Return String.Format(CultureInfo.InvariantCulture,
            "border:{0}pt solid {1};", BorderWidthPt, BorderColor)
    End Function

    Public Function TitleStyle() As String
        Return String.Format(CultureInfo.InvariantCulture,
            "background:{0};color:{1};font-family:'{2}';font-weight:bold;text-align:center;",
            TitleBackgroundColor, TitleFontColor, FontNameBold)
    End Function

    ' =====================================================================
    '  page.css yang ditulis ke folder output (dibangun dari setting di atas)
    ' =====================================================================
    Public Function BuildPageCss() As String
        Dim inv = CultureInfo.InvariantCulture
        Dim sb As New Text.StringBuilder()
        sb.AppendLine("@page { size: A4; margin: 0; }")
        sb.AppendLine("* { box-sizing: border-box; }")
        sb.AppendLine(String.Format(inv,
            "body {{ margin:0; font-family:'{0}'; font-size:{1}pt; color:{2}; -webkit-print-color-adjust:exact; print-color-adjust:exact; }}",
            FontNameNormal, FontSizePt, FontColor))
        sb.AppendLine(String.Format(inv,
            ".page {{ width:{0}mm; height:{1}mm; position:relative; overflow:hidden; page-break-after:always; }}",
            PageWidthMm, PageHeightMm))
        sb.AppendLine(String.Format(inv, ".hdr, .bdy, .ftr {{ position:absolute; left:0; width:{0}mm; }}", PageWidthMm))
        sb.AppendLine(".t   { position:absolute; white-space:nowrap; line-height:1; }")
        sb.AppendLine(".r   { position:absolute; }")
        sb.AppendLine(".box { " & BorderStyle() & " }")
        sb.AppendLine(".bar { " & TitleStyle() & " }")
        Return sb.ToString()
    End Function

End Module
