' =====================================================================
'  CondEvaluator.vb — evaluasi ekspresi <<if …>> di VB terhadap data (JSON)
'
'  Ekspresi ditulis di PDF seperti JavaScript sederhana:
'     riders.contains('FEC')            riders.includes("SOC")
'     riders.length > 0                 plan == 'A' && !(riders.contains('POC') || age >= 60)
'     not riders.contains('FEC') and footer.agentcode != ''
'  Nilai variabel = key di data (RiplayData.Build): riders → data("riders"); a.b.c = bersarang.
'  Method: .contains(x) / .includes(x) (array atau string, tanpa beda huruf besar-kecil), .length / .count,
'          .startsWith(x), .endsWith(x), .toLowerCase(), .toUpperCase(), .trim()
'  Operator: && and || or ! not  == = != <> < <= > >=  ( )   literal: 'teks' "teks" 123 true false null
'  Benar/salah: bool apa adanya; angka ≠ 0; string tidak kosong; array/objek tidak kosong; null/tidak ada = salah.
'  Variabel yang tidak ada di data → salah + dicatat (log) — tidak melempar error.
' =====================================================================
Imports System.Text.Json
Imports System.Text.RegularExpressions

Public Class CondEvaluator

    Private ReadOnly _data As JsonElement
    Private ReadOnly _log As Action(Of String)
    Private _toks As List(Of String)
    Private _pos As Integer

    Public Sub New(data As JsonElement, log As Action(Of String))
        _data = data : _log = log
    End Sub

    ''' <summary>Evaluasi ekspresi → True/False. Error sintaks / variabel tidak ada → False (dicatat).</summary>
    Public Function Eval(expr As String) As Boolean
        Try
            _toks = Tokenize(expr) : _pos = 0
            Dim v = ParseOr()
            If _pos < _toks.Count Then Throw New FormatException("sisa token '" & _toks(_pos) & "'")
            Return Truthy(v)
        Catch ex As Exception
            _log?.Invoke($"  ! <<if {expr}>>: {ex.Message} → dianggap salah")
            Return False
        End Try
    End Function

    ' ---------- tokenizer ----------
    Private Shared ReadOnly TokRx As New Regex(
        "\G(?:(?<str>'(?:[^'\\]|\\.)*'|""(?:[^""\\]|\\.)*"")|(?<num>\d+(?:\.\d+)?)|(?<id>[A-Za-z_$][\w$]*)|(?<op>&&|\|\||==|!=|<>|<=|>=|[!<>=()\.,]))",
        RegexOptions.Compiled)

    Private Shared Function Tokenize(expr As String) As List(Of String)
        Dim s = Regex.Replace(expr, "[‘’‚‛]", "'")
        s = Regex.Replace(s, "[“”„‟]", """")
        Dim res As New List(Of String)
        Dim i = 0
        While i < s.Length
            If Char.IsWhiteSpace(s(i)) Then i += 1 : Continue While
            Dim m = TokRx.Match(s, i)
            If Not m.Success Then Throw New FormatException("karakter tak dikenal di '" & s.Substring(i) & "'")
            res.Add(m.Value)
            i = m.Index + m.Length
        End While
        Return res
    End Function

    ' ---------- parser (recursive descent) ----------
    Private Function Peek() As String
        Return If(_pos < _toks.Count, _toks(_pos), Nothing)
    End Function
    Private Function Take() As String
        Dim t = Peek() : _pos += 1 : Return t
    End Function
    Private Function IsOp(t As String, ParamArray ops As String()) As Boolean
        Return t IsNot Nothing AndAlso ops.Any(Function(o) String.Equals(o, t, StringComparison.OrdinalIgnoreCase))
    End Function

    Private Function ParseOr() As Object
        Dim v = ParseAnd()
        While IsOp(Peek(), "||", "or")
            Take()
            Dim r = ParseAnd()
            v = Truthy(v) OrElse Truthy(r)
        End While
        Return v
    End Function

    Private Function ParseAnd() As Object
        Dim v = ParseNot()
        While IsOp(Peek(), "&&", "and")
            Take()
            Dim r = ParseNot()
            v = Truthy(v) AndAlso Truthy(r)
        End While
        Return v
    End Function

    Private Function ParseNot() As Object
        If IsOp(Peek(), "!", "not") Then
            Take()
            Return Not Truthy(ParseNot())
        End If
        Return ParseCmp()
    End Function

    Private Function ParseCmp() As Object
        Dim l = ParsePrim()
        Dim op = Peek()
        If Not IsOp(op, "==", "=", "!=", "<>", "<", "<=", ">", ">=") Then Return l
        Take()
        Dim r = ParsePrim()
        Select Case op
            Case "==", "=" : Return Equal(l, r)
            Case "!=", "<>" : Return Not Equal(l, r)
        End Select
        Dim a = ToNum(l), b = ToNum(r)
        If Double.IsNaN(a) OrElse Double.IsNaN(b) Then
            Dim c = String.Compare(ToStr(l), ToStr(r), StringComparison.OrdinalIgnoreCase)
            Return If(op = "<", c < 0, If(op = "<=", c <= 0, If(op = ">", c > 0, c >= 0)))
        End If
        Return If(op = "<", a < b, If(op = "<=", a <= b, If(op = ">", a > b, a >= b)))
    End Function

    Private Function ParsePrim() As Object
        Dim t = Take()
        If t Is Nothing Then Throw New FormatException("ekspresi belum lengkap")
        If t = "(" Then
            Dim v = ParseOr()
            If Take() <> ")" Then Throw New FormatException("kurung tutup hilang")
            Return ParseMembers(v)
        End If
        If t(0) = "'"c OrElse t(0) = """"c Then Return ParseMembers(Unquote(t))
        Dim num As Double
        If Double.TryParse(t, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, num) Then Return num
        Select Case t.ToLowerInvariant()
            Case "true" : Return True
            Case "false" : Return False
            Case "null", "undefined", "nothing" : Return Nothing
        End Select
        If Not Regex.IsMatch(t, "^[A-Za-z_$][\w$]*$") Then Throw New FormatException("token tak terduga '" & t & "'")
        ' variabel: ambil dari data lalu turuni .field / .method(...)
        Dim val As Object = Nothing
        Dim el As JsonElement
        If _data.ValueKind = JsonValueKind.Object AndAlso _data.TryGetProperty(t, el) Then
            val = FromJson(el)
        Else
            _log?.Invoke($"  ! variabel '{t}' tidak ada di data → null")
        End If
        Return ParseMembers(val)
    End Function

    ' .field  |  .method(args)
    Private Function ParseMembers(v As Object) As Object
        While Peek() = "."
            Take()
            Dim name = Take()
            If name Is Nothing Then Throw New FormatException("nama setelah '.' hilang")
            If Peek() = "(" Then
                Take()
                Dim args As New List(Of Object)
                If Peek() <> ")" Then
                    args.Add(ParseOr())
                    While Peek() = ","
                        Take() : args.Add(ParseOr())
                    End While
                End If
                If Take() <> ")" Then Throw New FormatException("kurung tutup hilang setelah " & name)
                v = CallMethod(v, name, args)
            Else
                v = Member(v, name)
            End If
        End While
        Return v
    End Function

    ' ---------- nilai ----------
    Private Shared Function FromJson(el As JsonElement) As Object
        Select Case el.ValueKind
            Case JsonValueKind.String : Return el.GetString()
            Case JsonValueKind.Number : Return el.GetDouble()
            Case JsonValueKind.True : Return True
            Case JsonValueKind.False : Return False
            Case JsonValueKind.Null, JsonValueKind.Undefined : Return Nothing
            Case Else : Return el                      ' array / objek tetap JsonElement
        End Select
    End Function

    Private Function Member(v As Object, name As String) As Object
        Dim lname = name.ToLowerInvariant()
        If TypeOf v Is JsonElement Then
            Dim el = DirectCast(v, JsonElement)
            If el.ValueKind = JsonValueKind.Array AndAlso (lname = "length" OrElse lname = "count") Then Return CDbl(el.GetArrayLength())
            If el.ValueKind = JsonValueKind.Object Then
                For Each p In el.EnumerateObject()           ' toleran huruf besar-kecil
                    If String.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase) Then Return FromJson(p.Value)
                Next
            End If
        ElseIf TypeOf v Is String AndAlso (lname = "length" OrElse lname = "count") Then
            Return CDbl(DirectCast(v, String).Length)
        End If
        _log?.Invoke($"  ! field '{name}' tidak ada → null")
        Return Nothing
    End Function

    Private Shared Function CallMethod(v As Object, name As String, args As List(Of Object)) As Object
        Dim a0 = If(args.Count > 0, ToStr(args(0)), "")
        Select Case name.ToLowerInvariant()
            Case "contains", "includes", "has"
                If TypeOf v Is JsonElement Then
                    Dim el = DirectCast(v, JsonElement)
                    If el.ValueKind = JsonValueKind.Array Then
                        Return el.EnumerateArray().Any(Function(x) String.Equals(ToStr(FromJson(x)), a0, StringComparison.OrdinalIgnoreCase))
                    End If
                    Return False
                End If
                If v Is Nothing Then Return False
                Return ToStr(v).IndexOf(a0, StringComparison.OrdinalIgnoreCase) >= 0
            Case "startswith" : Return v IsNot Nothing AndAlso ToStr(v).StartsWith(a0, StringComparison.OrdinalIgnoreCase)
            Case "endswith" : Return v IsNot Nothing AndAlso ToStr(v).EndsWith(a0, StringComparison.OrdinalIgnoreCase)
            Case "tolowercase", "tolower" : Return ToStr(v).ToLowerInvariant()
            Case "touppercase", "toupper" : Return ToStr(v).ToUpperInvariant()
            Case "trim" : Return ToStr(v).Trim()
            Case "length", "count", "size"
                If TypeOf v Is JsonElement AndAlso DirectCast(v, JsonElement).ValueKind = JsonValueKind.Array Then Return CDbl(DirectCast(v, JsonElement).GetArrayLength())
                Return CDbl(ToStr(v).Length)
            Case Else
                Throw New FormatException("method '" & name & "' tidak dikenal")
        End Select
    End Function

    Public Shared Function Truthy(v As Object) As Boolean
        If v Is Nothing Then Return False
        If TypeOf v Is Boolean Then Return CBool(v)
        If TypeOf v Is Double Then Return CDbl(v) <> 0
        If TypeOf v Is String Then Return DirectCast(v, String).Length > 0
        If TypeOf v Is JsonElement Then
            Dim el = DirectCast(v, JsonElement)
            If el.ValueKind = JsonValueKind.Array Then Return el.GetArrayLength() > 0
            If el.ValueKind = JsonValueKind.Object Then Return el.EnumerateObject().Any()
            Return Truthy(FromJson(el))
        End If
        Return True
    End Function

    Private Shared Function Equal(l As Object, r As Object) As Boolean
        If l Is Nothing OrElse r Is Nothing Then Return l Is Nothing AndAlso r Is Nothing
        Dim a = ToNum(l), b = ToNum(r)
        If Not Double.IsNaN(a) AndAlso Not Double.IsNaN(b) Then Return a = b
        Return String.Equals(ToStr(l), ToStr(r), StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function ToNum(v As Object) As Double
        If TypeOf v Is Double Then Return CDbl(v)
        If TypeOf v Is Boolean Then Return If(CBool(v), 1, 0)
        Dim d As Double
        If TypeOf v Is String AndAlso Double.TryParse(DirectCast(v, String), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, d) Then Return d
        Return Double.NaN
    End Function

    Public Shared Function ToStr(v As Object) As String
        If v Is Nothing Then Return ""
        If TypeOf v Is Double Then Return CDbl(v).ToString(Globalization.CultureInfo.InvariantCulture)
        If TypeOf v Is Boolean Then Return If(CBool(v), "true", "false")
        If TypeOf v Is JsonElement Then
            Dim el = DirectCast(v, JsonElement)
            Return If(el.ValueKind = JsonValueKind.String, el.GetString(), el.GetRawText())
        End If
        Return v.ToString()
    End Function

    Private Shared Function Unquote(t As String) As String
        Dim inner = t.Substring(1, t.Length - 2)
        Return Regex.Replace(inner, "\\(.)", "$1")
    End Function

End Class
