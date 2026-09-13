' =====================================================================
'  Models.vb — model data hasil ekstraksi PDF
'  Semua koordinat sudah dalam mm dengan origin kiri-ATAS (seperti HTML).
' =====================================================================

''' <summary>Satu kata dari PDF.</summary>
Public Class WordItem
    Public Text As String
    Public XMm As Double          ' kiri
    Public WidthMm As Double
End Class

''' <summary>Potongan teks dengan style seragam pada satu baris (1 div di HTML).</summary>
Public Class TextRun
    Public XMm As Double          ' kiri kata pertama
    Public BaselineMm As Double   ' baseline (dari atas halaman)
    Public WidthMm As Double      ' kiri kata pertama s/d kanan kata terakhir
    Public SizePt As Double
    Public PdfFontName As String
    Public IsBold As Boolean
    Public IsItalic As Boolean
    Public ColorHex As String     ' warna asli di PDF, mis. "#111111"
    Public IsSuper As Boolean     ' superscript (diset SemanticHtmlWriter saat menyusun baris)
    Public Words As New List(Of WordItem)

    Public ReadOnly Property Text As String
        Get
            Return String.Join(" ", Words.Select(Function(w) w.Text))
        End Get
    End Property
End Class

Public Enum ShapeKind
    FillRect      ' kotak isi (bar judul, background sel tabel, garis tabel tipis)
    StrokeRect    ' kotak border (header, footer, kolom)
    Path          ' bentuk bebas (logo) → SVG
End Enum

''' <summary>Kotak / garis / path dari PDF.</summary>
Public Class Shape
    Public Kind As ShapeKind
    Public XMm As Double
    Public YMm As Double
    Public WidthMm As Double
    Public HeightMm As Double
    Public FillHex As String
    Public StrokeHex As String
    Public StrokeWidthPt As Double
    Public SvgPathMm As String    ' hanya untuk Kind = Path; koordinat absolut halaman (mm)
End Class

Public Class PageModel
    Public Number As Integer
    Public WidthMm As Double
    Public HeightMm As Double
    Public Runs As New List(Of TextRun)
    Public Shapes As New List(Of Shape)
End Class

''' <summary>Bagian halaman (header / body / footer). Koordinat item RELATIF terhadap TopMm.</summary>
Public Class Region
    Public Name As String
    Public TopMm As Double
    Public HeightMm As Double
    Public Runs As New List(Of TextRun)
    Public Shapes As New List(Of Shape)
End Class
