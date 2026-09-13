' =====================================================================
'  DirectiveProcessor.vb — blok kondisional dari penanda (directive) di PDF
'
'  Di PDF: teks merah, SATU BARIS SENDIRI, rata kiri (sejajar dengan blok yang dibungkus):
'     <<if riders.contains('FEC')>>            ← boleh juga: <<start: if riders.contains('FEC') >>
'     … isi yang hanya tampil kalau kondisi benar (paragraf, list, tabel, bar) …
'     <<endif>>                                ← boleh juga: <<end>>  /  <<end: if riders.contains('FEC') >>
'
'  Di HTML body, dua baris <p> penanda itu diganti menjadi:
'     <div class="cond" data-if="riders.contains('FEC')" style="margin-top:…">
'     …
'     </div>
'  Kondisi dievaluasi di browser oleh bind.js (applyData) dengan data dari VB
'  (nama variabel = key di data; .contains(x) = array/string mengandung x).
'
'  Syarat: penanda awal & akhir harus sejajar — tag HTML di antaranya seimbang
'  (mis. keduanya di luar <ol>, bukan satu di dalam <li> dan satu di luar).
'  Kalau tidak, penanda dibiarkan sebagai teks merah dan dicatat di log.
'
'  BARIS TABEL BERULANG — di PDF, satu baris tabel berisi placeholder <<array.field>>:
'     | <<funds.fundname>> | <<funds.funddesc>> | <<funds.fundrisk>> |
'  Di HTML baris itu menjadi (format sel tetap, teks merah/italic placeholder dibuang):
'     <tr data-each="funds"><td>{{fundname}}</td><td>{{funddesc}}</td><td>{{fundrisk}}</td></tr>
'  bind.js meng-clone baris per item data["funds"] dan mengganti {{field}} dengan nilainya
'  ({{$no}} = nomor urut). Teks lain di sel boleh ikut: "Rp <<funds.amount>>" → "Rp {{amount}}".
'
'  PAGE BREAK — di PDF, teks merah satu baris sendiri:
'     <<Page Break>>                 (juga: <<page-break>> / <<pagebreak>>)
'  Di HTML body baris <p> itu diganti  <div class="page-break"></div>  → paginate.js
'  (Generate Riplay / --assemble) memulai halaman baru di titik itu, jadi di AllPages.pdf
'  isi setelahnya ada di halaman berikutnya. Teks penandanya sendiri tidak ikut dicetak.
' =====================================================================
Imports System.Net
Imports System.Text.RegularExpressions

Public Class DirectiveProcessor

    ' <<if EXPR>>  |  <<start: if EXPR>>  |  <<start: EXPR>>
    Private Shared ReadOnly StartRx As New Regex("^<<\s*(?:start\b\s*:?\s*(?:if\b)?|if\b)\s*(?<expr>.+?)\s*>>$", RegexOptions.IgnoreCase)
    ' <<endif>>  |  <<end>>  |  <<end if>>  |  <<end: if EXPR>>  |  <</if>>
    Private Shared ReadOnly EndRx As New Regex("^<<\s*(?:end\b\s*:?\s*(?:if\b.*?)?|endif\b.*?|/\s*if\b.*?)\s*>>$", RegexOptions.IgnoreCase)
    ' <<Page Break>>  |  <<page-break>>  |  <<pagebreak>>
    Private Shared ReadOnly PageBreakRx As New Regex("^<<\s*page\s*[-_]?\s*break\s*>>$", RegexOptions.IgnoreCase)
    ' satu baris HTML = satu paragraf (writer menulis 1 <p> per baris; </p> boleh tidak ada)
    Private Shared ReadOnly ParaRx As New Regex("^(?<ind>\s*)<p\b(?<attr>[^>]*)>(?<inner>.*?)(?:</p>)?\s*$", RegexOptions.IgnoreCase)
    Private Shared ReadOnly MarginRx As New Regex("margin-top:\s*([\d.]+)mm", RegexOptions.IgnoreCase)
    Private Shared ReadOnly TagRx As New Regex("<(?<close>/?)(?<name>[a-z][a-z0-9]*)\b[^>]*?(?<self>/?)>", RegexOptions.IgnoreCase)
    Private Shared ReadOnly VoidTags As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From
        {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"}
    ' <<funds.fundname>> — placeholder field dari array (baris tabel berulang)
    Private Shared ReadOnly FieldRx As New Regex("<<\s*(?<arr>[A-Za-z_]\w*)\.(?<fld>[A-Za-z_$][\w$]*)\s*>>")
    ' satu baris HTML = satu <tr> (writer menulis 1 baris tabel per baris)
    Private Shared ReadOnly RowRx As New Regex("^(?<ind>\s*)<tr\b(?<attr>[^>]*)>(?<cells>.*)</tr>\s*$", RegexOptions.IgnoreCase)
    Private Shared ReadOnly CellRx As New Regex("<(?<tag>t[dh])\b(?<attr>[^>]*)>(?<inner>.*?)</\k<tag>>", RegexOptions.IgnoreCase)

    Private Class Mark
        Public LineIdx As Integer
        Public IsStart As Boolean
        Public Expr As String
        Public Indent As String
        Public MarginTop As String   ' "4.23" atau Nothing
    End Class

    ' ------------------------------------------------------------------
    '  Deteksi teks penanda (dipakai juga oleh SemanticHtmlWriter supaya
    '  baris penanda tidak digabung ke paragraf sebelah/berikutnya)
    ' ------------------------------------------------------------------
    Public Shared Function IsStart(text As String, ByRef expr As String) As Boolean
        Dim m = StartRx.Match(Normalize(text))
        expr = If(m.Success, m.Groups("expr").Value, Nothing)
        Return m.Success
    End Function

    Public Shared Function IsEnd(text As String) As Boolean
        Return EndRx.IsMatch(Normalize(text))
    End Function

    Public Shared Function IsPageBreak(text As String) As Boolean
        Return PageBreakRx.IsMatch(Normalize(text))
    End Function

    Public Shared Function IsDirective(text As String) As Boolean
        Dim e As String = Nothing
        Return IsStart(text, e) OrElse IsEnd(text) OrElse IsPageBreak(text)
    End Function

    ''' <summary>Rapikan teks dari PDF: kutip keriting → lurus, spasi ganda → satu.</summary>
    Private Shared Function Normalize(text As String) As String
        If text Is Nothing Then Return ""
        Dim t = Regex.Replace(text, "[‘’‚‛]", "'")      ' ‘ ’ ‚ ‛ → '
        t = Regex.Replace(t, "[“”„‟]", """""")          ' “ ” „ ‟ → "
        Return Regex.Replace(t, "\s+", " ").Trim()
    End Function

    ' ------------------------------------------------------------------
    '  Proses HTML body:
    '    <p>penanda</p> … <p>penanda</p>      →  <div class="cond" data-if="…"> … </div>
    '    <tr><td><<funds.fundname>></td>…     →  <tr data-each="funds"><td>{{fundname}}</td>…
    ' ------------------------------------------------------------------
    Public Shared Function Apply(html As String, log As Action(Of String)) As String
        If html Is Nothing OrElse html.IndexOf("&lt;&lt;", StringComparison.Ordinal) < 0 AndAlso html.IndexOf("<<", StringComparison.Ordinal) < 0 Then
            Return html
        End If
        Dim lines = html.Split(ControlChars.Lf)
        ApplyIfBlocks(lines, log)
        ApplyEachRows(lines, log)
        ApplyPageBreaks(lines, log)
        Return String.Join(ControlChars.Lf, lines)
    End Function

    ' ------------------------------------------------------------------
    '  <p>&lt;&lt;Page Break&gt;&gt;</p>  →  <div class="page-break"></div>
    '  (dikenali paginate.js: pindah halaman; elemen tidak ikut dicetak)
    ' ------------------------------------------------------------------
    Private Shared Sub ApplyPageBreaks(lines As String(), log As Action(Of String))
        For i = 0 To lines.Length - 1
            Dim m = ParaRx.Match(lines(i))
            If Not m.Success Then Continue For
            Dim text = WebUtility.HtmlDecode(Regex.Replace(m.Groups("inner").Value, "<[^>]+>", ""))
            If Not IsPageBreak(text) Then Continue For
            lines(i) = m.Groups("ind").Value & "<div class=""page-break""></div>"
            log?.Invoke($"  page break: baris {i + 1}")
        Next
    End Sub

    ' ------------------------------------------------------------------
    '  Blok kondisional <<if …>> … <<endif>>
    ' ------------------------------------------------------------------
    Private Shared Sub ApplyIfBlocks(lines As String(), log As Action(Of String))
        Dim marks As New List(Of Mark)
        For i = 0 To lines.Length - 1
            Dim m = ParaRx.Match(lines(i))
            If Not m.Success Then Continue For
            Dim text = WebUtility.HtmlDecode(Regex.Replace(m.Groups("inner").Value, "<[^>]+>", ""))
            Dim expr As String = Nothing
            If IsStart(text, expr) Then
                Dim mm = MarginRx.Match(m.Groups("attr").Value)
                marks.Add(New Mark With {.LineIdx = i, .IsStart = True, .Expr = expr, .Indent = m.Groups("ind").Value,
                                         .MarginTop = If(mm.Success, mm.Groups(1).Value, Nothing)})
            ElseIf IsEnd(text) Then
                marks.Add(New Mark With {.LineIdx = i, .IsStart = False, .Indent = m.Groups("ind").Value})
            End If
        Next
        If marks.Count = 0 Then Return

        ' pasangkan if/end (boleh bersarang)
        Dim stack As New Stack(Of Mark)
        Dim pairs As New List(Of Tuple(Of Mark, Mark))
        For Each mk In marks
            If mk.IsStart Then
                stack.Push(mk)
            ElseIf stack.Count = 0 Then
                log?.Invoke($"  ! <<end>> tanpa <<if>> (baris {mk.LineIdx + 1}) diabaikan")
            Else
                pairs.Add(Tuple.Create(stack.Pop(), mk))
            End If
        Next
        For Each s In stack
            log?.Invoke($"  ! <<if {s.Expr}>> tanpa <<end>> (baris {s.LineIdx + 1}) dibiarkan sebagai teks")
        Next

        For Each p In pairs
            Dim s = p.Item1, e = p.Item2
            Dim between = String.Join(ControlChars.Lf, lines, s.LineIdx + 1, e.LineIdx - s.LineIdx - 1)
            If Not IsBalanced(between) Then
                log?.Invoke($"  ! <<if {s.Expr}>> (baris {s.LineIdx + 1}–{e.LineIdx + 1}) tidak sejajar dengan <<end>> → dibiarkan sebagai teks")
                Continue For
            End If
            Dim style = If(s.MarginTop IsNot Nothing, $" style=""margin-top:{s.MarginTop}mm""", "")
            lines(s.LineIdx) = $"{s.Indent}<div class=""cond"" data-if=""{AttrEncode(s.Expr)}""{style}>"
            lines(e.LineIdx) = $"{e.Indent}</div>"
            log?.Invoke($"  blok kondisional: if {s.Expr}  (baris {s.LineIdx + 1}–{e.LineIdx + 1})")
        Next
    End Sub

    ' ------------------------------------------------------------------
    '  Baris tabel berulang: sel berisi <<array.field>> → <tr data-each="array"> + {{field}}
    ' ------------------------------------------------------------------
    Private Shared Sub ApplyEachRows(lines As String(), log As Action(Of String))
        For i = 0 To lines.Length - 1
            Dim m = RowRx.Match(lines(i))
            If Not m.Success OrElse m.Groups("attr").Value.Contains("data-each") Then Continue For

            Dim arrName As String = Nothing
            Dim lineNo = i + 1
            Dim cells = CellRx.Replace(m.Groups("cells").Value,
                Function(cm As Match) As String
                    Dim text = CellText(cm.Groups("inner").Value)
                    Dim fields = FieldRx.Matches(text)
                    If fields.Count = 0 Then Return cm.Value
                    For Each f As Match In fields
                        If arrName Is Nothing Then
                            arrName = f.Groups("arr").Value
                        ElseIf f.Groups("arr").Value <> arrName Then
                            log?.Invoke($"  ! baris {lineNo}: <<{f.Groups("arr").Value}.…>> beda array dengan <<{arrName}.…>>, dianggap field {arrName}")
                        End If
                    Next
                    ' format sel (style/background) tetap; isi diganti teks polos + {{field}} (warna merah/italic placeholder dibuang)
                    Dim newText = FieldRx.Replace(text, Function(f) "{{" & f.Groups("fld").Value & "}}")
                    Return "<" & cm.Groups("tag").Value & cm.Groups("attr").Value & ">" & WebUtility.HtmlEncode(newText) & "</" & cm.Groups("tag").Value & ">"
                End Function)
            If arrName Is Nothing Then Continue For

            lines(i) = $"{m.Groups("ind").Value}<tr data-each=""{AttrEncode(arrName)}""{m.Groups("attr").Value}>{cells}</tr>"
            log?.Invoke($"  baris berulang: each {arrName}  (baris {lineNo})")
        Next
    End Sub

    ''' <summary>Teks polos isi sel: &lt;br&gt; → spasi, tag dibuang, entity di-decode, kutip/spasi dirapikan.</summary>
    Private Shared Function CellText(inner As String) As String
        Dim t = Regex.Replace(inner, "<br\s*/?>", " ", RegexOptions.IgnoreCase)
        t = Regex.Replace(t, "<[^>]+>", "")
        Return Normalize(WebUtility.HtmlDecode(t))
    End Function

    ''' <summary>Semua tag pembuka di potongan HTML ini tertutup di potongan yang sama?</summary>
    Private Shared Function IsBalanced(html As String) As Boolean
        Dim depth = 0
        For Each m As Match In TagRx.Matches(Regex.Replace(html, "<!--.*?-->", "", RegexOptions.Singleline))
            If VoidTags.Contains(m.Groups("name").Value) OrElse m.Groups("self").Value = "/" Then Continue For
            If m.Groups("close").Value = "/" Then
                depth -= 1
                If depth < 0 Then Return False
            Else
                depth += 1
            End If
        Next
        Return depth = 0
    End Function

    ' hanya karakter yang wajib di-escape di nilai atribut, supaya data-if tetap terbaca
    Private Shared Function AttrEncode(s As String) As String
        Return s.Replace("&", "&amp;").Replace("""", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;")
    End Function

End Class
