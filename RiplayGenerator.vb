' =====================================================================
'  RiplayGenerator.vb — tombol "Generate Riplay": PDF template → AllPages.html berisi data
'
'  Langkah:
'    1. Kumpulkan variabel VB yang dikirim ke HTML → RiplayData.Build() (sementara hardcode).
'    2. Pecah PDF → header.html, footer.html, bodyN.html, PageN.html  (Generator.Run).
'    3. Tulis semua variabel ke SATU file data.js (RiplayData.Write: window.riplayData = {...});
'       PageN.html / AllBody.html / dokumen kerja PageAssembler memuatnya lewat <script src="data.js">.
'    4. Blok <<if …>> … <<endif>> (div.cond data-if) dievaluasi bind.js saat halaman dibuka:
'       riders = {"SOC"} → blok if riders.contains('FEC') disembunyikan.
'    5. Gabungkan body1.html … bodyN.html → AllBody.html (satu .bdy panjang;
'       tiap halaman sumber dipisah <div class="page-break"> bila BreakBetweenSourcePages).
'    6. header.html + footer.html + AllBody.html → AllPages.html (PageAssembler + WebView2;
'       data diambil dari data.js, blok yang tersembunyi tidak ikut).
'    7. AllPages.html → AllPages.pdf (PdfExporter: WebView2 PrintToPdfAsync, A4, margin 0).
' =====================================================================
Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports Microsoft.Web.WebView2.WinForms

Public Class RiplayGenerator

    Public Const AllBodyFile As String = "AllBody.html"
    Public Const AllPagesPdf As String = "AllPages.pdf"

    Private Shared ReadOnly Utf8NoBom As New UTF8Encoding(False)
    ' isi di dalam <div class="bdy" …> … </div> (greedy → sampai </div> terakhir)
    Private Shared ReadOnly BdyRx As New Regex("^\s*<div class=""bdy""[^>]*>\s*(.*)</div>\s*$", RegexOptions.Singleline)
    Private Shared ReadOnly BdyTopRx As New Regex("<div class=""bdy""[^>]*top:\s*([\d.]+)mm", RegexOptions.IgnoreCase)

    ''' <summary>Jalankan langkah 1–7. Mengembalikan folder output.</summary>
    Public Shared Async Function RunAsync(pdfPath As String, web As WebView2, log As Action(Of String)) As Task(Of String)
        ' ---- 1. data: semua variabel VB di satu Dictionary (RiplayData.Build) ----
        Dim data = RiplayData.Build()
        log("Langkah 1: kumpulkan variabel → " & JsonSerializer.Serialize(data))

        ' ---- 2. pecah PDF ----
        log("Langkah 2: pecah PDF → header.html, footer.html, bodyN.html, PageN.html")
        Dim outDir = Await Task.Run(Function() Generator.Run(pdfPath, log))

        ' ---- 3. tulis variabel ke satu file data.js; semua HTML tinggal memuatnya ----
        Dim dataJs = RiplayData.Write(outDir, data)
        log("Langkah 3: " & dataJs & " ← window.riplayData = {…}  (dimuat PageN.html / AllBody.html / AllPages)")

        ' ---- 4. cocokkan otomatis: semua <<variabel>> di HTML <-> key di data.js ----
        log("Langkah 4: variabel di header/footer/body <-> data.js (bind.js mengganti saat halaman dibuka / dipecah)")
        Dim missing = RiplayData.ScanPlaceholders(outDir, data, log)
        If missing.Count > 0 Then log("  -> belum ada di RiplayData.Build(): " & String.Join(", ", missing))

        ' ---- 5. gabung body ----
        Dim bodies = Directory.GetFiles(outDir, "body*.html").
            Select(Function(f) New With {.Path = f, .N = BodyNumber(f)}).
            Where(Function(x) x.N > 0).OrderBy(Function(x) x.N).Select(Function(x) x.Path).ToList()
        If bodies.Count = 0 Then Throw New InvalidOperationException("Tidak ada bodyN.html di " & outDir)
        Dim allBodyPath = MergeBodies(outDir, bodies)
        log($"Langkah 5: {bodies.Count} body → {AllBodyFile}" &
            If(GlobalSettings.BreakBetweenSourcePages, " (tiap halaman sumber mulai di halaman baru)", " (mengalir)"))

        ' ---- 6. header + footer + AllBody → AllPages ----
        log("Langkah 6: header.html + footer.html + AllBody.html + data.js → AllPages.html")
        ' Data = Nothing → PageAssembler memuat data.js dari folder (satu sumber data untuk semua HTML)
        Dim asm As New PageAssembler With {.BodyFile = AllBodyFile}
        Dim allPages = Await asm.AssembleAsync(outDir, web, log)
        log("  ✓ " & allPages)

        ' ---- 7. AllPages.html → PDF (WebView2 PrintToPdf) ----
        Dim pdfOut = Path.Combine(outDir, AllPagesPdf)
        log($"Langkah 7: AllPages.html → {AllPagesPdf}")
        Await PdfExporter.ExportAsync(allPages, pdfOut, web, log)
        log("  ✓ " & pdfOut)
        Return outDir
    End Function

    Private Shared Function BodyNumber(file As String) As Integer
        Dim m = Regex.Match(Path.GetFileName(file), "^body(\d+)\.html$", RegexOptions.IgnoreCase)
        Return If(m.Success, Integer.Parse(m.Groups(1).Value), 0)
    End Function

    ''' <summary>
    ''' Langkah 5: isi semua bodyN.html digabung ke satu &lt;div class="bdy"&gt; di AllBody.html.
    ''' File ini bisa dibuka di browser (halaman memanjang, data.js diterapkan) dan menjadi
    ''' input PageAssembler (BodyFile = AllBody.html).
    ''' </summary>
    Public Shared Function MergeBodies(outDir As String, bodyFiles As IList(Of String)) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html><head><meta charset=""utf-8""><title>All Body</title><link rel=""stylesheet"" href=""" & GlobalSettings.CssFileName & """>")
        sb.AppendLine("<script src=""bind.js""></script>")
        ' preview: halaman memanjang mengikuti isi; garis putus-putus = batas halaman sumber
        sb.AppendLine("<style>.page{height:auto;min-height:297mm;overflow:visible}.bdy{position:relative;height:auto}" &
                      ".page-break{border-top:1px dashed #999;margin:4mm 0}</style></head><body>")
        sb.AppendLine("<div class=""page"">")

        Dim topMm As String = Nothing
        Dim first = True
        For Each f In bodyFiles
            Dim html = File.ReadAllText(f, Encoding.UTF8)
            If topMm Is Nothing Then
                Dim mt = BdyTopRx.Match(html)
                topMm = If(mt.Success, mt.Groups(1).Value, "0")
                sb.AppendLine($"<div class=""bdy"" style=""top:{topMm}mm"">")
            End If
            Dim m = BdyRx.Match(html)
            Dim inner = If(m.Success, m.Groups(1).Value, html).TrimEnd()
            If Not first AndAlso GlobalSettings.BreakBetweenSourcePages Then sb.AppendLine("<div class=""page-break""></div>")
            sb.AppendLine($"<!-- {Path.GetFileName(f)} -->")
            sb.AppendLine(inner)
            first = False
        Next

        sb.AppendLine("</div>")   ' .bdy
        sb.AppendLine("</div>")   ' .page
        sb.AppendLine("<script src=""data.js""></script>")
        sb.AppendLine("<script>applyData(window.riplayData === undefined ? null : window.riplayData);</script>")
        sb.AppendLine("</body></html>")

        Dim outPath = Path.Combine(outDir, AllBodyFile)
        File.WriteAllText(outPath, sb.ToString(), Utf8NoBom)
        Return outPath
    End Function

End Class
