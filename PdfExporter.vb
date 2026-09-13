' =====================================================================
'  PdfExporter.vb — Tahap 7: HTML (AllPages.html) → PDF
'
'  Memakai WebView2 (CoreWebView2.PrintToPdfAsync) yang sudah ada di form,
'  jadi tidak butuh Chrome/printer eksternal. Ukuran halaman & margin 0 ikut
'  GlobalSettings (A4); page.css sudah punya @page{size:A4;margin:0} dan
'  .page{page-break-after:always} sehingga 1 div.page = 1 halaman PDF.
' =====================================================================
Imports System.IO
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms

Public Class PdfExporter

    ''' <summary>Render file HTML di WebView2 lalu cetak ke PDF. Mengembalikan path PDF.</summary>
    Public Shared Async Function ExportAsync(htmlPath As String, pdfPath As String, web As WebView2, log As Action(Of String)) As Task(Of String)
        If Not File.Exists(htmlPath) Then Throw New FileNotFoundException("File HTML tidak ditemukan", htmlPath)
        Await web.EnsureCoreWebView2Async()

        ' ---------- muat HTML ----------
        Dim tcs As New TaskCompletionSource(Of Boolean)
        Dim handler As EventHandler(Of CoreWebView2NavigationCompletedEventArgs) =
            Sub(s, e) tcs.TrySetResult(e.IsSuccess)
        AddHandler web.CoreWebView2.NavigationCompleted, handler
        web.CoreWebView2.Navigate("file:///" & htmlPath.Replace("\", "/"))
        Dim ok = Await tcs.Task
        RemoveHandler web.CoreWebView2.NavigationCompleted, handler
        If Not ok Then Throw New InvalidOperationException("Gagal memuat " & Path.GetFileName(htmlPath) & " di WebView2")

        ' tunggu font siap supaya hasil cetak sama dengan tampilan
        For i = 1 To 50
            Dim st = Await web.CoreWebView2.ExecuteScriptAsync("document.fonts.status")
            If st.Contains("loaded") Then Exit For
            Await Task.Delay(100)
        Next

        ' ---------- setting cetak: A4, margin 0, tanpa header/footer browser, background ikut ----------
        Dim ps = web.CoreWebView2.Environment.CreatePrintSettings()
        ps.Orientation = CoreWebView2PrintOrientation.Portrait
        ps.ScaleFactor = 1
        ps.PageWidth = GlobalSettings.PageWidthMm / 25.4      ' inci
        ps.PageHeight = GlobalSettings.PageHeightMm / 25.4
        ps.MarginTop = 0 : ps.MarginBottom = 0 : ps.MarginLeft = 0 : ps.MarginRight = 0
        ps.ShouldPrintBackgrounds = True
        ps.ShouldPrintHeaderAndFooter = False
        ps.ShouldPrintSelectionOnly = False

        If File.Exists(pdfPath) Then File.Delete(pdfPath)
        Dim printed = Await web.CoreWebView2.PrintToPdfAsync(pdfPath, ps)
        If Not printed OrElse Not File.Exists(pdfPath) Then Throw New InvalidOperationException("PrintToPdf gagal: " & pdfPath)

        log?.Invoke($"PDF: {Path.GetFileName(pdfPath)} ({New FileInfo(pdfPath).Length \ 1024} KB)")
        Return pdfPath
    End Function

End Class
