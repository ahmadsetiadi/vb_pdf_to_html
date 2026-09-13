' =====================================================================
'  Generator.vb — orkestrasi: PDF → extract → split → tulis HTML
' =====================================================================
Imports System.IO

Public Class Generator

    ''' <summary>Folder output: &lt;folder PDF&gt;\Output\&lt;nama pdf&gt;\</summary>
    Public Shared Function OutputDirFor(pdfPath As String) As String
        Return Path.Combine(Path.GetDirectoryName(pdfPath), GlobalSettings.OutputFolder,
                            Path.GetFileNameWithoutExtension(pdfPath))
    End Function

    Public Shared Function Run(pdfPath As String, log As Action(Of String)) As String
        If Not File.Exists(pdfPath) Then Throw New FileNotFoundException("File PDF tidak ditemukan", pdfPath)

        Dim outDir = OutputDirFor(pdfPath)
        log($"Output: {outDir}")

        ' bersihkan output lama
        If Directory.Exists(outDir) Then
            For Each f In Directory.GetFiles(outDir, "*.html") : File.Delete(f) : Next
            For Each f In Directory.GetFiles(outDir, "*.css") : File.Delete(f) : Next
        End If
        Directory.CreateDirectory(outDir)
        File.Copy(pdfPath, Path.Combine(outDir, "source.pdf"), True)

        log("Tahap 2: ekstraksi PDF …")
        Dim pages = New PdfExtractor().Extract(pdfPath, log)

        log("Tahap 3: pisah header / body / footer …")
        Dim split = New RegionSplitter().Split(pages, log)

        log("Tahap 4: tulis HTML …")
        Dim files As List(Of String)
        If String.Equals(GlobalSettings.OutputMode, "positional", StringComparison.OrdinalIgnoreCase) Then
            files = New HtmlWriter().Write(outDir, pages, split, log)
        Else
            files = New SemanticHtmlWriter().Write(outDir, pages, split, log)
        End If
        For Each f In files
            log("  ✓ " & Path.GetFileName(f))
        Next

        log($"Selesai: {pages.Count} halaman → {files.Count} file.")
        Return outDir
    End Function

End Class
