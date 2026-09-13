' =====================================================================
'  PageAssembler.vb — Fase 2: satukan header.html + Page1.html (body panjang)
'  + footer.html menjadi AllPages.html yang sudah dipecah per halaman A4.
'
'  Pengukuran tinggi blok dilakukan browser (WebView2 tersembunyi) lewat
'  paginate.js; hasil DOM disimpan sebagai HTML statis.
' =====================================================================
Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports Microsoft.Web.WebView2.WinForms

Public Class PageAssembler

    Public Property HeaderFile As String = "header.html"
    Public Property FooterFile As String = "footer.html"
    Public Property BodyFile As String = "Page1.html"
    Public Property CssFile As String = "page.css"
    Public Property OutputFile As String = "AllPages.html"

    ''' <summary>
    ''' Data dari VB untuk template, mis. {"riders": {"FEC","SOC"}, "footer": {...}}. Semua directive yang masih ada di
    ''' header/footer/body (&lt;&lt;if&gt;&gt;, &lt;&lt;arr.field&gt;&gt;, &lt;&lt;var&gt;&gt;) dijalankan di VB
    ''' (DirectiveProcessor.Execute) sebelum dipecah; &lt;&lt;page&gt;&gt;/&lt;&lt;totalpages&gt;&gt; diisi paginate.js.
    ''' Nothing = pakai data.js di folder kalau ada; kalau tidak ada → mode template (semua blok tampil).
    ''' </summary>
    Public Property Data As IDictionary(Of String, Object)

    Private ReadOnly Utf8NoBom As New UTF8Encoding(False)

    ''' <summary>Gabungkan & pecah halaman. Mengembalikan path AllPages.html.</summary>
    Public Async Function AssembleAsync(folder As String, web As WebView2, log As Action(Of String)) As Task(Of String)
        Dim hdr = ReadBody(Path.Combine(folder, HeaderFile))
        Dim ftr = ReadBody(Path.Combine(folder, FooterFile))
        Dim bdy = ReadBody(Path.Combine(folder, BodyFile))
        log($"Baca {HeaderFile}, {FooterFile}, {BodyFile}")

        ' ---------- data: properti Data, kalau tidak ada → data.js di folder ----------
        Dim dataDict = Data
        If dataDict Is Nothing Then
            dataDict = RiplayData.Read(folder)
            If dataDict IsNot Nothing Then log("Data: dari " & RiplayData.FileName)
        End If
        If dataDict IsNot Nothing Then
            ' semua directive yang masih tersisa (header/footer: <<var>>; body yang diedit manual: <<if>>, <<arr.field>>)
            ' dijalankan di sini (VB). <<page>>/<<totalpages>> dibiarkan → diisi paginate.js per halaman.
            Dim dataEl = JsonSerializer.SerializeToElement(dataDict)
            hdr = DirectiveProcessor.Execute(hdr, dataEl, log)
            ftr = DirectiveProcessor.Execute(ftr, dataEl, log)
            bdy = DirectiveProcessor.Execute(bdy, dataEl, log)
        Else
            ' mode template: penanda <<if …>> → div.cond (bind.js menampilkan semua blok)
            bdy = DirectiveProcessor.Apply(bdy, log)
            log("Data: tidak ada → mode template (semua blok kondisional tampil, <<var>> dibiarkan)")
        End If

        ' ---------- dokumen kerja: css + template + script ----------
        Dim js = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "paginate.js"), Encoding.UTF8)
        Dim bindPath = Path.Combine(AppContext.BaseDirectory, "bind.js")
        Dim bind = If(File.Exists(bindPath), File.ReadAllText(bindPath, Encoding.UTF8), "")
        Dim work As New StringBuilder()
        work.AppendLine("<!DOCTYPE html><html><head><meta charset=""utf-8""><title>work</title>")
        work.AppendLine($"<link rel=""stylesheet"" href=""{CssFile}"">")
        work.AppendLine("<style>" & OverrideCss() & "</style></head><body>")
        work.AppendLine("<template id=""tplHeader"">" & hdr & "</template>")
        work.AppendLine("<template id=""tplFooter"">" & ftr & "</template>")
        work.AppendLine("<template id=""tplBody"">" & bdy & "</template>")
        work.AppendLine("<div id=""pages""></div>")
        If bind <> "" Then work.AppendLine("<script>" & bind & "</script>")   ' replacePlaceholders untuk <<page>>/<<totalpages>>
        work.AppendLine("<script>" & js & "</script>")
        work.AppendLine("</body></html>")
        Dim workPath = Path.Combine(folder, "_work.html")
        File.WriteAllText(workPath, work.ToString(), Utf8NoBom)

        ' ---------- muat di WebView2 ----------
        Await web.EnsureCoreWebView2Async()
        Dim tcs As New TaskCompletionSource(Of Boolean)
        Dim handler As EventHandler(Of Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs) =
            Sub(s, e) tcs.TrySetResult(e.IsSuccess)
        AddHandler web.CoreWebView2.NavigationCompleted, handler
        web.CoreWebView2.Navigate("file:///" & workPath.Replace("\", "/"))
        Dim ok = Await tcs.Task
        RemoveHandler web.CoreWebView2.NavigationCompleted, handler
        If Not ok Then Throw New InvalidOperationException("Gagal memuat dokumen kerja di WebView2")

        ' tunggu font siap supaya pengukuran tinggi akurat
        For i = 1 To 50
            Dim st = Await web.CoreWebView2.ExecuteScriptAsync("document.fonts.status")
            If st.Contains("loaded") Then Exit For
            Await Task.Delay(100)
        Next

        ' ---------- jalankan pemecah ----------
        log("Mengukur & memecah halaman …")
        Dim inv = Globalization.CultureInfo.InvariantCulture
        ' directive sudah dijalankan di VB → paginate() tanpa data (bind.js hanya mengisi <<page>>/<<totalpages>>)
        Dim call_ = String.Format(inv, "paginate({0}, {1})", GlobalSettings.HeaderBodyGapMm, GlobalSettings.BodyFooterGapMm)
        log($"Jarak header→body {GlobalSettings.HeaderBodyGapMm} mm, body→footer {GlobalSettings.BodyFooterGapMm} mm")
        Dim raw = Await web.CoreWebView2.ExecuteScriptAsync(call_)
        If raw = "null" OrElse String.IsNullOrEmpty(raw) Then Throw New InvalidOperationException("paginate() tidak mengembalikan hasil (lihat _work.html di browser untuk error)")
        Dim json = JsonSerializer.Deserialize(Of String)(raw)          ' hasil script = string JSON
        Dim doc = JsonDocument.Parse(json)
        Dim pagesCount = doc.RootElement.GetProperty("pages").GetInt32()
        Dim pagesHtml = doc.RootElement.GetProperty("html").GetString()

        ' ---------- tulis AllPages.html ----------
        Dim sb As New StringBuilder()
        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html><head><meta charset=""utf-8""><title>All Pages</title>")
        sb.AppendLine($"<link rel=""stylesheet"" href=""{CssFile}"">")
        sb.AppendLine("<style>" & OverrideCss() & "</style></head><body>")
        sb.AppendLine(pagesHtml)
        sb.AppendLine("</body></html>")
        Dim outPath = Path.Combine(folder, OutputFile)
        File.WriteAllText(outPath, sb.ToString(), Utf8NoBom)
        File.Delete(workPath)

        log($"{pagesCount} halaman → {OutputFile}")
        Return outPath
    End Function

    ' pastikan ukuran halaman A4 walau page.css diubah (mis. height:600mm untuk body panjang)
    Private Function OverrideCss() As String
        Return ".page{width:" & GlobalSettings.PageWidthMm.ToString(Globalization.CultureInfo.InvariantCulture) & "mm;height:" &
               GlobalSettings.PageHeightMm.ToString(Globalization.CultureInfo.InvariantCulture) & "mm;overflow:hidden}.bdy{overflow:hidden}" &
               "[hidden]{display:none!important}"
    End Function

    ''' <summary>Ambil isi di antara &lt;body&gt;…&lt;/body&gt; (kalau file utuh) atau seluruh file (kalau fragment). Komentar HTML & &lt;script&gt; dibuang.</summary>
    Private Function ReadBody(file As String) As String
        If Not IO.File.Exists(file) Then Throw New FileNotFoundException("File tidak ditemukan", file)
        Dim s = IO.File.ReadAllText(file, Encoding.UTF8)
        Dim m = Regex.Match(s, "<body[^>]*>(.*)</body>", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
        If m.Success Then s = m.Groups(1).Value
        s = Regex.Replace(s, "<!--.*?-->", "", RegexOptions.Singleline)
        s = Regex.Replace(s, "<script.*?</script>", "", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
        Return s
    End Function

End Class
