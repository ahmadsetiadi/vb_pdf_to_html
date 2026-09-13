' =====================================================================
'  RiplayData.vb — kumpulan variabel VB yang dikirim ke HTML (langkah 1 & 3 Generate Riplay)
'
'  Build()  : susun semua variabel di satu Dictionary (key = nama variabel yang ditulis di PDF).
'  Write()  : tulis Dictionary itu ke SATU file  Output\<nama>\data.js  berisi
'                 window.riplayData = { ...JSON rapi... };
'             File ini yang dimuat semua HTML (PageN.html, AllBody.html, dokumen kerja PageAssembler)
'             lewat <script src="data.js"> → bind.js (applyData) menerapkannya ke <<if …>> dan <<array.field>>.
'             Dipakai .js (bukan .json) karena browser menolak fetch() file JSON lokal (file://);
'             isinya tetap JSON murni di kanan tanda "=" sehingga mudah dibaca / diedit / dipakai program lain.
'
'  Ganti isi Build() dengan data dari aplikasi; HTML tidak perlu diubah.
' =====================================================================
Imports System.IO
Imports System.Text
Imports System.Text.Encodings.Web
Imports System.Text.Json
Imports System.Text.RegularExpressions

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

    ' ---------------------------------------------------------------------------------
    '  Pencocokan otomatis variabel HTML <-> data
    ' ---------------------------------------------------------------------------------
    ''' <summary>Placeholder &lt;&lt;nama&gt;&gt; / &lt;&lt;obj.field&gt;&gt; di HTML (bentuk ter-escape &amp;lt;&amp;lt;...&amp;gt;&amp;gt; maupun mentah).</summary>
    Private Shared ReadOnly PlaceholderRx As New Regex("(?:<<|&lt;&lt;)\s*([A-Za-z_$][\w$]*(?:\.[\w$]+)*)\s*(?:>>|&gt;&gt;)")
    ''' <summary>Placeholder yang diisi paginate.js (nomor halaman), bukan data.</summary>
    Private Shared ReadOnly PageVars As String() = {"page", "pageno", "totalpages", "total", "pages"}

    ''' <summary>
    ''' Cari semua &lt;&lt;variabel&gt;&gt; di header.html, footer.html, body*.html lalu cocokkan dengan data:
    ''' ada -> diganti bind.js saat halaman dibuka / dipecah; tidak ada -> dibiarkan (tetap merah).
    ''' Mengembalikan daftar path yang tidak ada di data.
    ''' </summary>
    Public Shared Function ScanPlaceholders(outDir As String, data As IDictionary(Of String, Object), log As Action(Of String)) As List(Of String)
        Dim files = New List(Of String) From {IO.Path.Combine(outDir, GlobalSettings.HeaderFileName), IO.Path.Combine(outDir, GlobalSettings.FooterFileName)}
        files.AddRange(Directory.GetFiles(outDir, "body*.html").OrderBy(Function(f) f, StringComparer.OrdinalIgnoreCase))
        Dim found As New SortedDictionary(Of String, SortedSet(Of String))(StringComparer.Ordinal)
        For Each f In files
            If Not File.Exists(f) Then Continue For
            For Each m As Match In PlaceholderRx.Matches(File.ReadAllText(f, Encoding.UTF8))
                Dim key = m.Groups(1).Value
                If Not found.ContainsKey(key) Then found(key) = New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
                found(key).Add(IO.Path.GetFileName(f))
            Next
        Next
        Dim json = JsonSerializer.SerializeToElement(data)
        Dim missing As New List(Of String)
        For Each kv In found
            Dim status As String
            If PageVars.Contains(kv.Key.ToLowerInvariant()) Then
                status = "nomor halaman (diisi paginate.js)"
            Else
                Dim v = Resolve(json, kv.Key)
                If Not v.HasValue Then
                    status = "TIDAK ADA di data.js -> dibiarkan" : missing.Add(kv.Key)
                ElseIf v.Value.ValueKind = JsonValueKind.Array OrElse v.Value.ValueKind = JsonValueKind.Object Then
                    status = "array/objek -> hanya untuk <<if>> / baris tabel, bukan teks"
                Else
                    status = "= " & v.Value.ToString()
                End If
            End If
            log($"  <<{kv.Key}>> [{String.Join(", ", kv.Value)}] {status}")
        Next
        If found.Count = 0 Then log("  (tidak ada <<variabel>> di HTML)")
        Return missing
    End Function

    Private Shared Function Resolve(root As JsonElement, path As String) As JsonElement?
        Dim cur = root
        For Each seg In path.Split("."c)
            If cur.ValueKind <> JsonValueKind.Object Then Return Nothing
            Dim nxt As JsonElement
            If Not cur.TryGetProperty(seg, nxt) Then Return Nothing
            cur = nxt
        Next
        Return cur
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
