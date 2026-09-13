' =====================================================================
'  RiplayGenerator.vb — tombol "Generate Riplay": PDF template → AllPages.html / AllPages.pdf berisi data
'
'  Langkah:
'    1. Kumpulkan variabel VB yang dikirim ke HTML → RiplayData.Build() (sementara hardcode).
'    2. Pecah PDF → header.html, footer.html, bodyN.html, PageN.html (Generator.Run).
'       HTML-nya STATIS persis seperti PDF: <<if>>, <<endif>>, <<Page Break>>, <<var.field>> tetap teks merah.
'    3. Tulis semua variabel ke SATU file data.js (RiplayData.Write: window.riplayData = {...}).
'    4. Jalankan semua directive dengan data (DirectiveProcessor.Execute, di VB):
'       gabungkan body1..N → satu teks, lalu <<if>> dievaluasi (blok dipertahankan/dibuang), baris tabel
'       <<arr.field>> di-clone per item, <<var>> diganti nilainya, <<Page Break>> → div.page-break
'       → AllBody.html tanpa directive sama sekali (hanya <<page>>/<<totalpages>> di footer yang menunggu langkah 5).
'    5. header.html + footer.html + AllBody.html → AllPages.html (PageAssembler + WebView2 paginate.js):
'       <<var>> di header/footer diganti, isi dipecah per halaman A4, <<page>>/<<totalpages>> diisi per halaman.
'    6. AllPages.html → AllPages.pdf (PdfExporter: WebView2 PrintToPdfAsync, A4, margin 0).
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

    ''' <summary>Jalankan langkah 1–6. Mengembalikan folder output.</summary>
    Public Shared Async Function RunAsync(pdfPath As String, web As WebView2, log As Action(Of String)) As Task(Of String)
        ' ---- 1. data: semua variabel VB di satu Dictionary (RiplayData.Build) ----
        Dim data = RiplayData.Build()
        log("Langkah 1: kumpulkan variabel → " & JsonSerializer.Serialize(data))

        ' ---- 2. pecah PDF → HTML statis (directive tetap teks merah) ----
        log("Langkah 2: pecah PDF → header.html, footer.html, bodyN.html, PageN.html (directive <<…>> dibiarkan seperti di PDF)")
        Dim outDir = Await Task.Run(Function() Generator.Run(pdfPath, log))

        ' ---- 3. tulis variabel ke satu file data.js ----
        Dim dataJs = RiplayData.Write(outDir, data)
        log("Langkah 3: " & dataJs & " ← window.riplayData = {…}")

        ' ---- 4. jalankan semua directive dengan data → AllBody.html tanpa <<…>> ----
        Dim bodies = Directory.GetFiles(outDir, "body*.html").
            Select(Function(f) New With {.Path = f, .N = BodyNumber(f)}).
            Where(Function(x) x.N > 0).OrderBy(Function(x) x.N).Select(Function(x) x.Path).ToList()
        If bodies.Count = 0 Then Throw New InvalidOperationException("Tidak ada bodyN.html di " & outDir)
        log($"Langkah 4: jalankan directive di {bodies.Count} body dengan data.js → {AllBodyFile}" &
            If(GlobalSettings.BreakBetweenSourcePages, " (tiap halaman sumber mulai di halaman baru)", " (mengalir; halaman baru hanya di <<Page Break>>)"))
        Dim json = JsonSerializer.SerializeToElement(data)
        Dim allBodyPath = MergeBodies(outDir, bodies, json, log)
        log("  ✓ " & allBodyPath)

        ' ---- 5. header + footer + AllBody → AllPages (dipecah per halaman, <<page>> diisi) ----
        log("Langkah 5: header.html + footer.html + AllBody.html → AllPages.html (pecah per A4, isi <<page>>/<<totalpages>>)")
        Dim asm As New PageAssembler With {.BodyFile = AllBodyFile, .Data = data}
        Dim allPages = Await asm.AssembleAsync(outDir, web, log)
        log("  ✓ " & allPages)

        ' ---- 6. AllPages.html → PDF (WebView2 PrintToPdf) ----
        Dim pdfOut = Path.Combine(outDir, AllPagesPdf)
        log($"Langkah 6: AllPages.html → {AllPagesPdf}")
        Await PdfExporter.ExportAsync(allPages, pdfOut, web, log)
        log("  ✓ " & pdfOut)
        Return outDir
    End Function

    Private Shared Function BodyNumber(file As String) As Integer
        Dim m = Regex.Match(Path.GetFileName(file), "^body(\d+)\.html$", RegexOptions.IgnoreCase)
        Return If(m.Success, Integer.Parse(m.Groups(1).Value), 0)
    End Function

    ''' <summary>
    ''' Langkah 4: isi semua bodyN.html digabung ke satu &lt;div class="bdy"&gt;, lalu SEMUA directive dijalankan
    ''' dengan data (DirectiveProcessor.Execute) → AllBody.html statis tanpa &lt;&lt;…&gt;&gt;.
    ''' File ini bisa dibuka di browser (halaman memanjang) dan menjadi input PageAssembler (BodyFile = AllBody.html).
    ''' data = Nothing → directive dibiarkan (mode template).
    ''' </summary>
    Public Shared Function MergeBodies(outDir As String, bodyFiles As IList(Of String), data As JsonElement?, log As Action(Of String)) As String
        ' ---- gabung mentah ----
        Dim body As New StringBuilder()
        Dim topMm As String = Nothing
        Dim first = True
        For Each f In bodyFiles
            Dim html = File.ReadAllText(f, Encoding.UTF8)
            If topMm Is Nothing Then
                Dim mt = BdyTopRx.Match(html)
                topMm = If(mt.Success, mt.Groups(1).Value, "0")
            End If
            Dim m = BdyRx.Match(html)
            Dim inner = If(m.Success, m.Groups(1).Value, html).TrimEnd()
            If Not first AndAlso GlobalSettings.BreakBetweenSourcePages Then body.AppendLine("<div class=""page-break""></div>")
            body.AppendLine($"<!-- {Path.GetFileName(f)} -->")
            body.AppendLine(inner)
            first = False
        Next

        ' ---- jalankan directive: <<if>>, <<arr.field>>, <<var>>, <<Page Break>> ----
        Dim inner_ = body.ToString()
        If data.HasValue Then
            inner_ = DirectiveProcessor.Execute(inner_, data.Value, log)
            ' sisa <<…>> = catatan template yang bukan variabel (ada spasi/tanda baca) atau variabel yang belum ada di data
            Dim plain = System.Net.WebUtility.HtmlDecode(Regex.Replace(Regex.Replace(inner_, "<br\s*/?>", " ", RegexOptions.IgnoreCase), "<[^>]+>", ""))
            Dim sisa = Regex.Matches(plain, "<<[^<>
]{1,80}>>").Cast(Of Match)().Select(Function(x) Regex.Replace(x.Value, "\s+", " ")).Distinct().ToList()
            If sisa.Count > 0 Then log($"  sisa {sisa.Count} macam <<…>> yang bukan directive/variabel: " &
                                       String.Join(", ", sisa.Take(12)) & If(sisa.Count > 12, ", …", ""))
        End If

        Dim sb As New StringBuilder()
        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html><head><meta charset=""utf-8""><title>All Body</title><link rel=""stylesheet"" href=""" & GlobalSettings.CssFileName & """>")
        ' preview: halaman memanjang mengikuti isi; garis putus-putus = page break
        sb.AppendLine("<style>.page{height:auto;min-height:297mm;overflow:visible}.bdy{position:relative;height:auto}" &
                      ".page-break{border-top:1px dashed #999;margin:4mm 0}</style></head><body>")
        sb.AppendLine("<div class=""page"">")
        sb.AppendLine($"<div class=""bdy"" style=""top:{topMm}mm"">")
        sb.Append(inner_)
        sb.AppendLine("</div>")   ' .bdy
        sb.AppendLine("</div>")   ' .page
        sb.AppendLine("</body></html>")

        Dim outPath = Path.Combine(outDir, AllBodyFile)
        File.WriteAllText(outPath, sb.ToString(), Utf8NoBom)
        Return outPath
    End Function

End Class
