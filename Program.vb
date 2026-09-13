Friend Module Program

    ''' <summary>
    ''' Tanpa argumen → buka UI.
    ''' "--ui <pdf>" → buka UI dan langsung import PDF tersebut.
    ''' "--assemble <folder> [--riders FEC,SOC]" → satukan HTML di folder itu jadi AllPages.html lalu keluar.
    '''     --riders = contoh variabel VB (String()) yang dikirim ke template sebagai data("riders");
    '''     tanpa --riders → pakai data.js di folder (kalau ada) / mode template.
    ''' "--riplay <pdf>" → Generate Riplay (6 langkah, lihat RiplayGenerator.vb) lalu keluar.
    ''' "--html2pdf <folder>" → HTML to PDF: folder HTML hasil langkah 2 (boleh diedit) → data.js → AllBody → AllPages → AllPages.pdf lalu keluar.
    ''' Dengan argumen path PDF → mode CLI: generate lalu keluar (exit code 0 = sukses, 1 = gagal).
    ''' </summary>
    <STAThread()>
    Friend Function Main(args As String()) As Integer
        Dim initialPdf As String = Nothing
        Dim assembleFolder As String = Nothing
        Dim assembleData As Dictionary(Of String, Object) = Nothing
        Dim riplayPdf As String = Nothing
        Dim htmlFolder As String = Nothing
        If args.Length >= 2 AndAlso args(0) = "--html2pdf" Then
            htmlFolder = args(1)
        ElseIf args.Length >= 2 AndAlso args(0) = "--ui" Then
            initialPdf = args(1)
        ElseIf args.Length >= 2 AndAlso args(0) = "--riplay" Then
            riplayPdf = args(1)
        ElseIf args.Length >= 2 AndAlso args(0) = "--assemble" Then
            assembleFolder = args(1)
            Dim i = 2
            While i + 1 < args.Length
                If args(i) = "--riders" Then
                    ' variabel VB biasa → masuk ke Data; di HTML jadi: riders.contains('FEC')
                    Dim riders As String() = args(i + 1).Split(","c).Select(Function(r) r.Trim()).Where(Function(r) r <> "").ToArray()
                    If assembleData Is Nothing Then assembleData = New Dictionary(Of String, Object)
                    assembleData("riders") = riders
                End If
                i += 2
            End While
        ElseIf args.Length > 0 Then
            Try
                Try : Console.OutputEncoding = Text.Encoding.UTF8 : Catch : End Try
                Dim dir = Generator.Run(args(0), Sub(m) Console.WriteLine(m))
                Console.WriteLine(dir)
                Return 0
            Catch ex As Exception
                Console.Error.WriteLine("ERROR: " & ex.Message)
                Return 1
            End Try
        End If

        Application.SetHighDpiMode(HighDpiMode.SystemAware)
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        ' mode CLI (--assemble / --riplay) tidak punya console → log juga ke file di folder output
        Dim logFile As String = Nothing
        If riplayPdf IsNot Nothing Then logFile = IO.Path.Combine(Generator.OutputDirFor(riplayPdf), "riplay.log")
        If assembleFolder IsNot Nothing Then logFile = IO.Path.Combine(assembleFolder, "riplay.log")
        If htmlFolder IsNot Nothing Then logFile = IO.Path.Combine(htmlFolder, "riplay.log")
        If logFile IsNot Nothing AndAlso IO.File.Exists(logFile) Then IO.File.Delete(logFile)
        Application.Run(New Form1 With {.InitialPdf = initialPdf, .AssembleFolder = assembleFolder, .AssembleData = assembleData,
                                        .RiplayPdf = riplayPdf, .HtmlFolder = htmlFolder, .LogFile = logFile})
        Return 0
    End Function

End Module
