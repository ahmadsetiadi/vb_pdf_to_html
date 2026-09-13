<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        lblPdf = New Label()
        txtPdf = New TextBox()
        btnBrowse = New Button()
        btnGenerate = New Button()
        btnOpenOutput = New Button()
        txtLog = New TextBox()
        lblHint = New Label()
        btnGeneratePdf = New Button()
        btnGenerateRiplay = New Button()
        btnHtmlToPdf = New Button()
        web = New Microsoft.Web.WebView2.WinForms.WebView2()
        CType(web, System.ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' btnGeneratePdf
        ' 
        btnGeneratePdf.Location = New Point(327, 42)
        btnGeneratePdf.Name = "btnGeneratePdf"
        btnGeneratePdf.Size = New Size(130, 25)
        btnGeneratePdf.Text = "Generate PDF"
        btnGeneratePdf.UseVisualStyleBackColor = True
        '
        ' btnHtmlToPdf (folder HTML hasil langkah 2, boleh diedit → data.js → AllBody → AllPages → AllPages.pdf)
        ' 
        btnHtmlToPdf.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnHtmlToPdf.Location = New Point(599, 42)
        btnHtmlToPdf.Name = "btnHtmlToPdf"
        btnHtmlToPdf.Size = New Size(112, 25)
        btnHtmlToPdf.TabIndex = 9
        btnHtmlToPdf.Text = "HTML to PDF"
        btnHtmlToPdf.UseVisualStyleBackColor = True
        ' 
        ' btnGenerateRiplay (PDF → HTML + data → AllBody.html → AllPages.html)
        '
        btnGenerateRiplay.Location = New Point(463, 42)
        btnGenerateRiplay.Name = "btnGenerateRiplay"
        btnGenerateRiplay.Size = New Size(130, 25)
        btnGenerateRiplay.Text = "Generate Riplay"
        btnGenerateRiplay.UseVisualStyleBackColor = True
        ' 
        ' web (WebView2 tersembunyi untuk mengukur & memecah halaman)
        ' 
        web.Location = New Point(-2000, -2000)
        web.Name = "web"
        web.Size = New Size(900, 1300)
        web.TabStop = False
        '
        ' lblPdf
        '
        lblPdf.AutoSize = True
        lblPdf.Location = New Point(12, 15)
        lblPdf.Name = "lblPdf"
        lblPdf.Size = New Size(52, 15)
        lblPdf.Text = "File PDF:"
        '
        ' txtPdf
        '
        txtPdf.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        txtPdf.Location = New Point(75, 12)
        txtPdf.Name = "txtPdf"
        txtPdf.ReadOnly = True
        txtPdf.Size = New Size(520, 23)
        '
        ' btnBrowse
        '
        btnBrowse.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnBrowse.Location = New Point(601, 11)
        btnBrowse.Name = "btnBrowse"
        btnBrowse.Size = New Size(110, 25)
        btnBrowse.Text = "Import PDF…"
        btnBrowse.UseVisualStyleBackColor = True
        '
        ' btnGenerate
        '
        btnGenerate.Enabled = False
        btnGenerate.Location = New Point(75, 42)
        btnGenerate.Name = "btnGenerate"
        btnGenerate.Size = New Size(110, 25)
        btnGenerate.Text = "Generate ulang"
        btnGenerate.UseVisualStyleBackColor = True
        '
        ' btnOpenOutput
        '
        btnOpenOutput.Enabled = False
        btnOpenOutput.Location = New Point(191, 42)
        btnOpenOutput.Name = "btnOpenOutput"
        btnOpenOutput.Size = New Size(130, 25)
        btnOpenOutput.Text = "Buka folder output"
        btnOpenOutput.UseVisualStyleBackColor = True
        '
        ' lblHint
        '
        lblHint.AutoSize = True
        lblHint.ForeColor = SystemColors.GrayText
        lblHint.Location = New Point(12, 78)
        lblHint.Name = "lblHint"
        lblHint.UseMnemonic = False
        lblHint.Text = "Import PDF (atau drag & drop ke jendela ini) → HTML otomatis di-generate ke folder Output\<nama pdf>\"
        '
        ' txtLog
        '
        txtLog.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        txtLog.Font = New Font("Consolas", 9F)
        txtLog.Location = New Point(12, 104)
        txtLog.Multiline = True
        txtLog.Name = "txtLog"
        txtLog.ReadOnly = True
        txtLog.ScrollBars = ScrollBars.Vertical
        txtLog.Size = New Size(699, 335)
        '
        ' Form1
        '
        AllowDrop = True
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(723, 451)
        Controls.Add(web)
        Controls.Add(btnGeneratePdf)
        Controls.Add(btnGenerateRiplay)
        Controls.Add(btnHtmlToPdf)
        Controls.Add(lblHint)
        Controls.Add(txtLog)
        Controls.Add(btnOpenOutput)
        Controls.Add(btnGenerate)
        Controls.Add(btnBrowse)
        Controls.Add(txtPdf)
        Controls.Add(lblPdf)
        MinimumSize = New Size(600, 350)
        Name = "Form1"
        StartPosition = FormStartPosition.CenterScreen
        Text = "Riplay PDF → HTML Generator (Fase 1)"
        CType(web, System.ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents lblPdf As Label
    Friend WithEvents txtPdf As TextBox
    Friend WithEvents btnBrowse As Button
    Friend WithEvents btnGenerate As Button
    Friend WithEvents btnOpenOutput As Button
    Friend WithEvents txtLog As TextBox
    Friend WithEvents lblHint As Label
    Friend WithEvents btnGeneratePdf As Button
    Friend WithEvents btnGenerateRiplay As Button
    Friend WithEvents btnHtmlToPdf As Button
    Friend WithEvents web As Microsoft.Web.WebView2.WinForms.WebView2
End Class
