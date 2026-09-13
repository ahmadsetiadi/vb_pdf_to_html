Imports System.IO

Public Class Form1

    Private outputDir As String

    ''' <summary>Buka form lalu langsung import PDF ini (dipakai mode "--ui <pdf>").</summary>
    Public Property InitialPdf As String

    ''' <summary>Mode "--assemble &lt;folder&gt;": langsung satukan HTML di folder ini lalu tutup.</summary>
    Public Property AssembleFolder As String

    ''' <summary>Data untuk template (mis. "riders" → String()). Nothing = pakai data.js di folder / mode template.</summary>
    Public Property AssembleData As IDictionary(Of String, Object)

    ''' <summary>Mode "--riplay &lt;pdf&gt;": jalankan Generate Riplay untuk PDF ini lalu tutup.</summary>
    Public Property RiplayPdf As String

    ''' <summary>Kalau diisi, semua baris log juga ditulis ke file ini (dipakai mode CLI yang tidak punya console).</summary>
    Public Property LogFile As String

    Private Async Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If Not String.IsNullOrEmpty(InitialPdf) Then ImportAndGenerate(InitialPdf)
        If Not String.IsNullOrEmpty(AssembleFolder) Then
            Dim ok = Await AssembleAsync(AssembleFolder)
            Environment.Exit(If(ok, 0, 1))
        End If
        If Not String.IsNullOrEmpty(RiplayPdf) Then
            Dim ok = Await GenerateRiplayAsync(RiplayPdf)
            Environment.Exit(If(ok, 0, 1))
        End If
    End Sub

    ' ---------- Generate Riplay: PDF → HTML + data → AllBody.html → AllPages.html (lihat RiplayGenerator.vb) ----------
    Private Async Sub btnGenerateRiplay_Click(sender As Object, e As EventArgs) Handles btnGenerateRiplay.Click
        Dim pdf = txtPdf.Text
        If pdf = "" OrElse Not File.Exists(pdf) Then
            Using dlg As New OpenFileDialog With {.Title = "Pilih PDF template Riplay", .Filter = "PDF (*.pdf)|*.pdf", .CheckFileExists = True}
                If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
                pdf = dlg.FileName
            End Using
        End If
        Await GenerateRiplayAsync(pdf)
    End Sub

    Private Async Function GenerateRiplayAsync(pdfPath As String) As Task(Of Boolean)
        txtPdf.Text = pdfPath
        txtLog.Clear()
        SetBusy(True)
        Log($"Generate Riplay: {pdfPath}")
        Try
            Dim dir = Await RiplayGenerator.RunAsync(pdfPath, web, AddressOf LogSafe)
            outputDir = dir
            btnOpenOutput.Enabled = True
            Log("")
            Log("Hasil akhir: AllPages.pdf (dan AllPages.html); preview per halaman: PageN.html / AllBody.html.")
            Return True
        Catch ex As Exception
            Log("ERROR: " & ex.Message)
            Log(ex.ToString())
            If String.IsNullOrEmpty(RiplayPdf) Then
                MessageBox.Show(Me, ex.Message, "Gagal Generate Riplay", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Return False
        Finally
            SetBusy(False)
        End Try
    End Function

    ' ---------- Fase 2: Generate PDF (langkah 1 = satukan HTML jadi AllPages.html) ----------
    Private Async Sub btnGeneratePdf_Click(sender As Object, e As EventArgs) Handles btnGeneratePdf.Click
        Dim start = If(outputDir IsNot Nothing AndAlso Directory.Exists(outputDir), outputDir,
                       If(txtPdf.Text <> "", Path.Combine(Path.GetDirectoryName(txtPdf.Text), GlobalSettings.OutputFolder), ""))
        Using dlg As New FolderBrowserDialog With {
            .Description = "Pilih folder berisi header.html, footer.html, Page1.html, page.css",
            .UseDescriptionForTitle = True,
            .InitialDirectory = start
        }
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await AssembleAsync(dlg.SelectedPath)
        End Using
    End Sub

    Private Async Function AssembleAsync(folder As String) As Task(Of Boolean)
        SetBusy(True)
        Log("")
        Log($"Generate PDF – langkah 1: satukan HTML di {folder}")
        Try
            Dim asm As New PageAssembler With {.Data = AssembleData}
            Dim outPath = Await asm.AssembleAsync(folder, web, AddressOf LogSafe)
            Log("  ✓ " & outPath)
            outputDir = folder
            btnOpenOutput.Enabled = True
            Return True
        Catch ex As Exception
            Log("ERROR: " & ex.Message)
            If String.IsNullOrEmpty(AssembleFolder) Then
                MessageBox.Show(Me, ex.Message, "Gagal menyatukan HTML", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Return False
        Finally
            SetBusy(False)
        End Try
    End Function

    ' ---------- Import PDF → langsung generate ----------
    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Using dlg As New OpenFileDialog With {
            .Title = "Pilih file PDF",
            .Filter = "PDF (*.pdf)|*.pdf",
            .CheckFileExists = True
        }
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                ImportAndGenerate(dlg.FileName)
            End If
        End Using
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        If txtPdf.Text <> "" Then ImportAndGenerate(txtPdf.Text)
    End Sub

    Private Sub btnOpenOutput_Click(sender As Object, e As EventArgs) Handles btnOpenOutput.Click
        If outputDir IsNot Nothing AndAlso Directory.Exists(outputDir) Then
            Process.Start(New ProcessStartInfo("explorer.exe", outputDir) With {.UseShellExecute = True})
        End If
    End Sub

    ' ---------- Drag & drop PDF ke form ----------
    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files = CType(e.Data.GetData(DataFormats.FileDrop), String())
            If files.Length = 1 AndAlso files(0).EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                e.Effect = DragDropEffects.Copy
            End If
        End If
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        Dim files = CType(e.Data.GetData(DataFormats.FileDrop), String())
        ImportAndGenerate(files(0))
    End Sub

    ' ---------- Proses ----------
    Private Async Sub ImportAndGenerate(pdfPath As String)
        txtPdf.Text = pdfPath
        txtLog.Clear()
        ' SetBusy(True)
        ' Log($"Import: {pdfPath}")

        ' Try
        '     Dim dir = Await Task.Run(Function() Generator.Run(pdfPath, AddressOf LogSafe))
        '     outputDir = dir
        '     btnOpenOutput.Enabled = True
        '     Log("")
        '     Log("Buka PageN.html di browser untuk cek hasil.")
        ' Catch ex As Exception
        '     Log("ERROR: " & ex.Message)
        '     MessageBox.Show(Me, ex.Message, "Gagal generate", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ' Finally
        '     SetBusy(False)
        ' End Try
    End Sub

    Private Sub SetBusy(busy As Boolean)
        UseWaitCursor = busy
        btnBrowse.Enabled = Not busy
        btnGeneratePdf.Enabled = Not busy
        btnGenerateRiplay.Enabled = Not busy
        btnGenerate.Enabled = Not busy AndAlso txtPdf.Text <> ""
    End Sub

    Private Sub Log(msg As String)
        txtLog.AppendText(msg & Environment.NewLine)
        If Not String.IsNullOrEmpty(LogFile) Then
            Try
                Directory.CreateDirectory(Path.GetDirectoryName(LogFile))
                File.AppendAllText(LogFile, msg & Environment.NewLine)
            Catch
            End Try
        End If
    End Sub

    Private Sub LogSafe(msg As String)
        If InvokeRequired Then
            BeginInvoke(Sub() Log(msg))
        Else
            Log(msg)
        End If
    End Sub

End Class
