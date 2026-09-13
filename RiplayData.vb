' =====================================================================
'  RiplayData.vb — kumpulan variabel VB yang dikirim ke HTML (langkah 1 & 3 Generate Riplay)
'
'  Build()  : susun semua variabel di satu Dictionary (key = nama variabel yang ditulis di PDF).
'  Write()  : tulis Dictionary itu ke SATU file  Output\<nama>\data.js  berisi
'                 window.riplayData = { ...JSON rapi... };
'             Isinya JSON murni di kanan tanda "=" (mudah dibaca / diedit / dipakai program lain).
'  Read()   : baca kembali file itu (dipakai PageAssembler / --assemble bila Data tidak diberikan).
'  Directive di HTML (<<if>>, <<arr.field>>, <<var>>) dijalankan di VB oleh DirectiveProcessor.Execute
'  dengan data ini (Generate Riplay langkah 4–5) → AllBody.html / AllPages.html statis tanpa <<…>>.
'
'  Ganti isi Build() dengan data dari aplikasi; HTML tidak perlu diubah.
' =====================================================================
Imports System.IO
Imports System.Text
Imports System.Text.Encodings.Web
Imports System.Text.Json

Public Class RiplayData

    ''' <summary>Nama file data di folder output.</summary>
    Public Const FileName As String = "data.js"

    Private Shared ReadOnly Utf8NoBom As New UTF8Encoding(False)
    Private Shared ReadOnly JsonOpts As New JsonSerializerOptions With {
        .WriteIndented = True,
        .Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping}   ' huruf non-ASCII ditulis apa adanya, bukan \uXXXX

    ''' <summary>
    ''' Langkah 1 — semua variabel untuk template. Sementara hardcode; nanti diisi dari aplikasi.
    ''' Key dictionary = nama variabel yang ditulis di PDF:
    '''   riders → array string  → blok  <<if riders.contains('FEC')>> … <<endif>>
    '''   funds  → array objek   → baris tabel  <<funds.fundname>> | <<funds.funddesc>> | <<funds.fundrisk>>
    ''' Nilai boleh String / Double / Boolean / String() / List / Dictionary / objek (anonymous atau Class)
    ''' — diserialisasi ke JSON; nama property objek = nama field di PDF.
    ''' </summary>
    Public Shared Function Build() As Dictionary(Of String, Object)
        Dim riders As String() = {"FEC"}                 ' contoh: {"FEC", "SOC", "POC"}

        ' array objek → tiap item = 1 baris tabel; property = field yang dipakai <<funds.xxx>>
        Dim funds = {
            New With {.fundname = "Manulife Dana Ekuitas", .funddesc = "dana ekuitas", .fundrisk = "Sedang"},
            New With {.fundname = "Manulife Dana Sejahtera", .funddesc = "dana sejahtera", .fundrisk = "Tinggi"},
            New With {.fundname = "Manulife Dana Umum", .funddesc = "dana umum", .fundrisk = "Rendah"}
        }

        Dim footer = New With {.agentname = "Budi", .agentcode = "A123", .proposalnumber = "P-001",
                       .printdate = "13/09/2026", .expireddate = "13/10/2026"}

        Return New Dictionary(Of String, Object) From {
            {"riders", riders},
            {"funds", funds},
            {"footer", footer}
        }
    End Function

    ''' <summary>JSON rapi (indent) dari data.</summary>
    Public Shared Function ToJson(data As IDictionary(Of String, Object)) As String
        Return JsonSerializer.Serialize(data, JsonOpts)
    End Function

    ''' <summary>Langkah 3 — tulis data ke Output\&lt;nama&gt;\data.js. Mengembalikan path file.</summary>
    Public Shared Function Write(outDir As String, data As IDictionary(Of String, Object)) As String
        Dim path = IO.Path.Combine(outDir, FileName)
        Dim sb As New StringBuilder()
        sb.AppendLine("// data.js — semua variabel dari VB (RiplayData.Build) untuk template HTML.")
        sb.AppendLine("// Ditulis ulang tiap Generate Riplay. Dimuat PageN.html / AllBody.html / PageAssembler lewat <script src=""data.js"">.")
        sb.AppendLine("// Boleh diedit lalu klik Generate PDF (atau --assemble) untuk menyusun ulang AllPages tanpa mengulang import.")
        sb.AppendLine("window.riplayData = " & ToJson(data) & ";")
        File.WriteAllText(path, sb.ToString(), Utf8NoBom)
        Return path
    End Function

    ''' <summary>Baca kembali data.js di folder (bagian JSON-nya). Nothing kalau file tidak ada.</summary>
    Public Shared Function Read(outDir As String) As Dictionary(Of String, Object)
        Dim path = IO.Path.Combine(outDir, FileName)
        If Not File.Exists(path) Then Return Nothing
        Dim txt = File.ReadAllText(path, Encoding.UTF8)
        Dim i = txt.IndexOf("=")
        Dim j = txt.LastIndexOf(";")
        If i < 0 OrElse j <= i Then Return Nothing
        Return JsonSerializer.Deserialize(Of Dictionary(Of String, Object))(txt.Substring(i + 1, j - i - 1))
    End Function

End Class
