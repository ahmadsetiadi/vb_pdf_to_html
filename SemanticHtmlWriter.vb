' =====================================================================
'  SemanticHtmlWriter.vb — Tahap 4 (mode "semantic")
'  Rekonstruksi struktur dari data posisi → HTML seperti riplay.html:
'    header  : .header.box  (logo svg + nama + produk)  flex space-between
'    body    : .bar, .cols/.col.box, <ol class="lst"> (marker + isi), <p>, <table class="tbl">
'    footer  : .footer.box  (.meta key:value, .disclaimer, .pageno)
'  Posisi vertikal dijaga lewat margin-top (selisih posisi asli), lebar lewat width mm.
' =====================================================================
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions

Public Class SemanticHtmlWriter

    Private ReadOnly inv As CultureInfo = CultureInfo.InvariantCulture
    Private ReadOnly Utf8NoBom As New UTF8Encoding(False)
    Private Const MmPerPt As Double = 25.4 / 72
    Private Const Ascent As Double = 0.905      ' Arial
    Private Const Descent As Double = 0.212
    Private pageW, pageH As Double
    Private log As Action(Of String)

    ' ------------------------------------------------------------------
    '  Struktur bantu
    ' ------------------------------------------------------------------
    Private Class Line
        Public Baseline, Left, Right, Size As Double
        Public Segs As New List(Of TextRun)
        Public Consumed As Boolean            ' sudah dipakai blok lain (mis. baris label:nilai)
        Public Top, Height As Double          ' line-box (mm, koordinat region); diisi RenderBlocks
        Public BlockHtml As String            ' bukan Nothing = "baris blok": html jadi (tabel/bar/kotak/kv/cols)
        Public ReadOnly Property IsBlock As Boolean
            Get
                Return BlockHtml IsNot Nothing
            End Get
        End Property
        Public ReadOnly Property AllBold As Boolean
            Get
                Dim main = Segs.Where(Function(s) Not s.IsSuper).ToList()
                Return main.Count > 0 AndAlso main.All(Function(s) s.IsBold)
            End Get
        End Property
        Public ReadOnly Property Text As String
            Get
                Return String.Join(" ", Segs.Select(Function(s) s.Text))
            End Get
        End Property
    End Class

    ''' <summary>Kotak konten: batas kiri/kanan teks & tinggi baris (mm).</summary>
    Private Class Ctx
        Public Left, Right, Top As Double
        Public Size As Double        ' pt dominan
        Public LineH As Double       ' mm
        Public PrevBottom As Double  ' posisi bawah blok terakhir (mm, koordinat region)
    End Class

    ' node hasil parsing list/paragraf
    Private Class Node
        Public Kind As String        ' "root" | "list" | "item" | "para" | "block"
        Public Children As New List(Of Node)
        Public Parent As Node
        Public MarkerX, ContentX As Double
        Public MarkerSeg As TextRun
        Public Lines As New List(Of Line)   ' para
        Public MarginTop As Double
        Public MarginLeft As Double  ' block
        Public Html As String        ' block (dengan placeholder __MT__/__ML__)
        Public Indent As Double      ' para: baris pertama menjorok dari kiri kontainer (mm)
        Public IsSpread As Boolean   ' para: segmen diposisikan absolut
        Public BaseX As Double       ' para: x kiri kontainer (untuk posisi absolut)
        Public Sub Add(n As Node)
            n.Parent = Me : Children.Add(n)
        End Sub
    End Class

    ' ==================================================================
    Public Function Write(outDir As String, pages As List(Of PageModel), split As SplitResult, logger As Action(Of String)) As List(Of String)
        log = logger
        Dim written As New List(Of String)
        Directory.CreateDirectory(outDir)
        pageW = pages(0).WidthMm : pageH = pages(0).HeightMm

        Dim cssPath = Path.Combine(outDir, GlobalSettings.CssFileName)
        File.WriteAllText(cssPath, BuildCss(), Utf8NoBom) : written.Add(cssPath)

        Dim headerHtml = RenderHeader(split.Header)
        Dim footerHtml = RenderFooter(split.Footer)
        Dim hp = Path.Combine(outDir, GlobalSettings.HeaderFileName)
        Dim fp = Path.Combine(outDir, GlobalSettings.FooterFileName)
        File.WriteAllText(hp, headerHtml, Utf8NoBom) : written.Add(hp)
        File.WriteAllText(fp, footerHtml, Utf8NoBom) : written.Add(fp)

        ' Langkah 2 = HTML statis persis seperti PDF: directive <<...>> tetap teks merah, tanpa bind.js/data.js.
        ' Directive dijalankan di Generate Riplay langkah 4 (DirectiveProcessor.Execute) saat digabung ke AllBody.html.

        ' PageN.html = BODY SAJA (tanpa header/footer) + link page.css, mulai HeaderBodyGapMm dari atas;
        ' margin-top blok pertama dibuang. Inilah file body per halaman (tidak ada bodyN.html terpisah).
        For i = 0 To pages.Count - 1
            Dim n = i + 1
            Dim bodyHtml = RenderBody(split.Bodies(i))               ' directive dibiarkan apa adanya
            Dim pp = Path.Combine(outDir, $"Page{n}.html")
            File.WriteAllText(pp, HtmlHead($"Page {n}", BodyOnlyCss()) & vbLf & "<div class=""page"">" & vbLf & bodyHtml & vbLf & "</div>" & vbLf & HtmlTail(), Utf8NoBom)
            written.Add(pp)
        Next
        Return written
    End Function

    Private Function HtmlHead(title As String, Optional extraCss As String = "") As String
        Return "<!DOCTYPE html>" & vbLf & "<html><head><meta charset=""utf-8""><title>" & WebUtility.HtmlEncode(title) &
               "</title><link rel=""stylesheet"" href=""" & GlobalSettings.CssFileName & """>" &
               If(extraCss <> "", "<style>" & extraCss & "</style>", "") & "</head><body>"
    End Function

    ''' <summary>
    ''' CSS preview PageN.html: hanya body — tinggi mengikuti isi, mulai HeaderBodyGapMm dari atas.
    ''' margin-top blok pertama (berlapis: flow > bar, ol > li, cols > col > bar, juga kolom ke-2) dibuang,
    ''' sama seperti yang dilakukan paginate.js (trimTop) di AllPages.
    ''' </summary>
    Private Function BodyOnlyCss() As String
        Dim sels As New List(Of String)
        Dim chain = ".bdy"
        For lvl = 1 To 8
            chain &= ">:first-child"
            sels.Add(chain)
            sels.Add(chain.Substring(0, chain.Length - ">:first-child".Length) & ">.col>:first-child")
        Next
        Return ".page{height:auto;min-height:0;overflow:visible}" &
               Fmt(".bdy{{position:relative;top:0!important;height:auto!important;padding-top:{0}mm}}", R2(GlobalSettings.HeaderBodyGapMm)) &
               String.Join(",", sels) & "{margin-top:0!important}"
    End Function

    Private Function HtmlTail() As String
        Return "</body></html>"
    End Function

    Private Function BuildCss() As String
        Dim sb As New StringBuilder(GlobalSettings.BuildPageCss())
        Dim fi = GlobalSettings.ResolveFont(False, True)
        sb.AppendLine(Fmt(".it {{ font-family:'{0}',{1}; font-style:{2}; }}", fi.FontName, GlobalSettings.FontFallback, fi.FontStyle))
        sb.AppendLine("[hidden] { display:none !important; }")   ' blok kondisional non-aktif (bind.js)
        sb.AppendLine("p { margin:0; text-align:justify; }")
        sb.AppendLine("p.sub { font-weight:bold; }")
        sb.AppendLine(".lst { list-style:none; margin:0; padding:0; }")
        sb.AppendLine(".lst > li { display:flex; }")
        sb.AppendLine(".lst > li > .m { flex:0 0 auto; }")
        sb.AppendLine(".lst > li > .b { flex:1; min-width:0; }")
        sb.AppendLine(".header { display:flex; justify-content:space-between; align-items:center; }")
        sb.AppendLine(".logo { display:flex; align-items:center; }")
        sb.AppendLine(".logo .name { font-weight:bold; line-height:1; }")
        sb.AppendLine(".product { font-weight:bold; }")
        sb.AppendLine(".cols { display:flex; align-items:stretch; }")
        sb.AppendLine(".col { overflow:hidden; }")
        sb.AppendLine("table.tbl { border-collapse:collapse; table-layout:fixed; }")
        sb.AppendLine(Fmt(".tbl td, .tbl th {{ border:{0}pt solid {1}; text-align:center; vertical-align:middle; padding:0 1mm; }}",
                        GlobalSettings.BorderWidthPt, GlobalSettings.BorderColor))
        sb.AppendLine(".footer .foot-row { display:flex; justify-content:space-between; align-items:center; }")
        sb.AppendLine(".meta > div { display:flex; white-space:nowrap; }")
        sb.AppendLine(".meta .k, .meta .c { flex:0 0 auto; }")
        sb.AppendLine(".disclaimer { font-weight:bold; text-align:center; white-space:nowrap; }")
        sb.AppendLine(".pageno { text-align:right; }")
        Return sb.ToString()
    End Function

    ' ==================================================================
    '  HEADER
    ' ==================================================================
    Private Function RenderHeader(r As Region) As String
        Dim sb As New StringBuilder()
        sb.AppendLine(Fmt("<div class=""hdr"" style=""top:{0}mm;height:{1}mm"">", R2(r.TopMm), R2(r.HeightMm)))

        Dim box = r.Shapes.Where(Function(s) s.Kind = ShapeKind.StrokeRect).OrderByDescending(Function(s) s.WidthMm).FirstOrDefault()
        Dim lines = BuildLines(r.Runs)
        Dim paths = r.Shapes.Where(Function(s) s.Kind = ShapeKind.Path).ToList()

        Dim bL = If(box IsNot Nothing, box.XMm, 0), bT = If(box IsNot Nothing, box.YMm, 0)
        Dim bR = If(box IsNot Nothing, box.XMm + box.WidthMm, pageW), bB = If(box IsNot Nothing, box.YMm + box.HeightMm, r.HeightMm)

        Dim leftLines = lines.Where(Function(l) l.Left < pageW / 2).OrderBy(Function(l) l.Left).ToList()
        Dim rightLines = lines.Where(Function(l) l.Left >= pageW / 2).ToList()

        Dim leftMost = Math.Min(If(paths.Any(), paths.Min(Function(p) p.XMm), Double.MaxValue),
                                If(leftLines.Any(), leftLines.Min(Function(l) l.Left), Double.MaxValue))
        Dim rightMost = If(rightLines.Any(), rightLines.Max(Function(l) l.Right), bR)

        sb.AppendLine(Fmt("<div class=""header{0}"" style=""margin:{1}mm {2}mm 0 {3}mm;height:{4}mm;padding:0 {5}mm 0 {6}mm"">",
                        If(box IsNot Nothing, " box", ""), R2(bT), R2(pageW - bR), R2(bL), R2(bB - bT),
                        R2(Math.Max(0, bR - rightMost)), R2(Math.Max(0, leftMost - bL))))

        ' ----- logo: svg + teks di sebelahnya -----
        sb.Append("  <div class=""logo""")
        If paths.Any() AndAlso leftLines.Any() Then
            Dim svgRight = paths.Max(Function(p) p.XMm + p.WidthMm)
            sb.Append(Fmt(" style=""gap:{0}mm""", R2(Math.Max(0, leftLines(0).Left - svgRight))))
        End If
        sb.AppendLine(">")
        If paths.Any() Then sb.AppendLine("    " & RenderSvgGroup(paths, r.TopMm))
        For Each l In leftLines
            sb.AppendLine(Fmt("    <span class=""name"" style=""{0}"">{1}</span>", GlobalSettings.TextStyle(l.AllBold, False, l.Size), InlineHtml(l, l.Size, l.AllBold)))
        Next
        sb.AppendLine("  </div>")

        ' ----- produk (kanan) -----
        For Each l In rightLines
            sb.AppendLine(Fmt("  <div class=""product"" style=""{0}"">{1}</div>", GlobalSettings.TextStyle(l.AllBold, False, l.Size), InlineHtml(l, l.Size, l.AllBold)))
        Next

        sb.AppendLine("</div>")
        sb.Append("</div>")
        Return sb.ToString()
    End Function

    ' gabung beberapa path (logo) jadi 1 svg dengan viewBox = bounding box-nya
    Private Function RenderSvgGroup(paths As List(Of Shape), regionTop As Double) As String
        Dim x0 = paths.Min(Function(p) p.XMm), y0 = paths.Min(Function(p) p.YMm) + regionTop
        Dim x1 = paths.Max(Function(p) p.XMm + p.WidthMm), y1 = paths.Max(Function(p) p.YMm + p.HeightMm) + regionTop
        Dim sb As New StringBuilder()
        sb.Append(Fmt("<svg width=""{0}mm"" height=""{1}mm"" viewBox=""{2} {3} {0} {1}"" style=""display:block"">", R2(x1 - x0), R2(y1 - y0), R2(x0), R2(y0)))
        For Each p In paths
            sb.Append(Fmt("<path fill=""{0}"" d=""{1}""/>", If(p.FillHex, "none"), p.SvgPathMm))
        Next
        sb.Append("</svg>")
        Return sb.ToString()
    End Function

    ' ==================================================================
    '  FOOTER
    ' ==================================================================
    Private Function RenderFooter(r As Region) As String
        Dim sb As New StringBuilder()
        sb.AppendLine(Fmt("<div class=""ftr"" style=""top:{0}mm;height:{1}mm"">", R2(r.TopMm), R2(r.HeightMm)))

        Dim strokes = r.Shapes.Where(Function(s) s.Kind = ShapeKind.StrokeRect).OrderByDescending(Function(s) s.WidthMm).ToList()
        Dim box = strokes.FirstOrDefault()
        Dim inner = strokes.Skip(1).ToList()

        Dim bL = If(box IsNot Nothing, box.XMm, 0), bT = If(box IsNot Nothing, box.YMm, 0)
        Dim bR = If(box IsNot Nothing, box.XMm + box.WidthMm, pageW), bB = If(box IsNot Nothing, box.YMm + box.HeightMm, r.HeightMm)

        ' partisi run: yang di dalam kotak disclaimer dipisah dulu
        Dim discBox = inner.FirstOrDefault()
        Dim discRuns = If(discBox IsNot Nothing, r.Runs.Where(Function(x) RunInside(x, discBox)).ToList(), New List(Of TextRun))
        Dim lines = BuildLines(r.Runs.Except(discRuns).ToList())
        Dim discLines = BuildLines(discRuns)

        ' baris meta = ada kata ":" di dalamnya
        Dim metaLines = lines.Where(Function(l) l.Segs.Any(Function(sg) sg.Words.Any(Function(w) w.Text = ":" OrElse (w.Text.Length > 1 AndAlso w.Text.EndsWith(":"))))).ToList()
        Dim rest = lines.Except(metaLines).OrderBy(Function(l) l.Baseline).ToList()

        Dim size = If(metaLines.Any(), metaLines(0).Size, If(lines.Any(), lines(0).Size, GlobalSettings.FontSizePt))
        Dim pitch As Double = LinePitch(metaLines, size)
        Dim metaTop = If(metaLines.Any(), LineTop(metaLines.Min(Function(l) l.Baseline), size, pitch), bT)
        Dim metaBottom = If(metaLines.Any(), LineTop(metaLines.Max(Function(l) l.Baseline), size, pitch) + pitch, bT)
        Dim leftMost = If(lines.Any(), lines.Min(Function(l) l.Left), bL)
        Dim rightMost = If(lines.Any(), lines.Max(Function(l) l.Right), bR)

        sb.AppendLine(Fmt("<div class=""footer{0}"" style=""margin:{1}mm {2}mm 0 {3}mm;height:{4}mm;padding:{5}mm {6}mm 0 {7}mm"">",
                        If(box IsNot Nothing, " box", ""), R2(bT), R2(pageW - bR), R2(bL), R2(bB - bT),
                        R2(metaTop - bT), R2(Math.Max(0, bR - rightMost)), R2(Math.Max(0, leftMost - bL))))
        sb.AppendLine(Fmt("  <div class=""foot-row"" style=""height:{0}mm"">", R2(metaBottom - metaTop)))

        ' ----- meta key : value -----
        If metaLines.Any() Then
            sb.AppendLine(Fmt("    <div class=""meta"" style=""font-size:{0}pt;line-height:{1}mm"">", R2(size), R2(pitch)))
            For Each l In metaLines.OrderBy(Function(x) x.Baseline)
                Dim keySegs As List(Of TextRun) = Nothing, colon As TextRun = Nothing, valSegs As List(Of TextRun) = Nothing
                SplitAtColon(l, keySegs, colon, valSegs)
                Dim keyX = If(keySegs.Any(), keySegs(0).XMm, l.Left)
                Dim valX = If(valSegs.Any(), valSegs(0).XMm, colon.XMm + colon.WidthMm)
                sb.AppendLine(Fmt("      <div><span class=""k"" style=""width:{0}mm"">{1}</span><span class=""c"" style=""width:{2}mm"">:</span><span class=""v"">{3}</span></div>",
                                R2(colon.XMm - keyX), InlineSegs(keySegs, size, False), R2(valX - colon.XMm), InlineSegs(valSegs, size, False)))
            Next
            sb.AppendLine("    </div>")
        End If

        ' ----- disclaimer (kotak) -----
        If discBox IsNot Nothing AndAlso discLines.Any() Then
            Dim dl = discLines(0)
            Dim padX = (discBox.WidthMm - (dl.Right - dl.Left)) / 2
            Dim padY = (discBox.HeightMm - dl.Size * MmPerPt * 1.117) / 2
            sb.AppendLine(Fmt("    <div class=""disclaimer"" style=""border:{0}pt solid {1};padding:{2}mm {3}mm;margin-right:{4}mm;{5}"">{6}</div>",
                            GlobalSettings.BorderWidthPt, GlobalSettings.BorderColor, R2(Math.Max(0, padY)), R2(Math.Max(0, padX)),
                            R2(Math.Max(0, rightMost - (discBox.XMm + discBox.WidthMm))),
                            GlobalSettings.TextStyle(dl.AllBold, False, dl.Size), InlineHtml(dl, dl.Size, dl.AllBold)))
        End If
        sb.AppendLine("  </div>")

        ' ----- sisa (nomor halaman) -----
        For Each l In rest
            Dim lp = LinePitch(New List(Of Line) From {l}, l.Size)
            sb.AppendLine(Fmt("  <div class=""pageno"" style=""margin-top:{0}mm;font-size:{1}pt;line-height:{2}mm"">{3}</div>",
                            R2(LineTop(l.Baseline, l.Size, lp) - metaBottom), R2(l.Size), R2(lp), InlineHtml(l, l.Size, False)))
        Next

        sb.AppendLine("</div>")
        sb.Append("</div>")
        Return sb.ToString()
    End Function

    ' ==================================================================
    '  BODY
    ' ==================================================================
    Private Function RenderBody(r As Region) As String
        Dim sb As New StringBuilder()
        sb.AppendLine(Fmt("<div class=""bdy"" style=""top:{0}mm;height:{1}mm"">", R2(r.TopMm), R2(r.HeightMm)))

        Dim runs = r.Runs.ToList()
        Dim shapes = r.Shapes.ToList()

        ' kolom = kotak stroke tinggi (>= 50% region) & tidak selebar halaman, >= 2 berdampingan
        Dim colBoxes = shapes.Where(Function(s) s.Kind = ShapeKind.StrokeRect AndAlso
                                              s.HeightMm >= 0.5 * r.HeightMm AndAlso
                                              s.WidthMm <= 0.7 * pageW).OrderBy(Function(s) s.XMm).ToList()

        If colBoxes.Count >= 2 Then
            Dim top = colBoxes.Min(Function(b) b.YMm)
            Dim bottom = colBoxes.Max(Function(b) b.YMm + b.HeightMm)
            Dim h = bottom - top
            Dim used As New HashSet(Of TextRun)
            For Each b In colBoxes
                Dim box = b
                For Each x In runs.Where(Function(rn) RunInside(rn, box)) : used.Add(x) : Next
            Next
            ' teks di luar kotak kolom: di atas kolom (mis. catatan, <<Page Break>>) dan di bawahnya
            ' dirender sebagai flow biasa; yang di antara kolom diabaikan
            Dim outside = runs.Where(Function(x) Not used.Contains(x)).ToList()
            Dim above = outside.Where(Function(x) RunCenterY(x) < top).ToList()
            Dim below = outside.Where(Function(x) RunCenterY(x) > bottom).ToList()
            Dim outShapes = shapes.Where(Function(s) Not colBoxes.Contains(s) AndAlso Not colBoxes.Any(Function(b) ShapeInside(s, b))).ToList()
            Dim prevBottom As Double = 0
            If above.Count > 0 Then
                prevBottom = RenderBlocks(sb, above, outShapes.Where(Function(s) s.YMm + s.HeightMm / 2 < top).ToList(), 0, pageW, 0)
            End If

            sb.AppendLine(Fmt("<div class=""cols"" style=""margin-top:{0}mm;margin-left:{1}mm;height:{2}mm"">", R2(top - prevBottom), R2(colBoxes(0).XMm), R2(h)))
            For ci = 0 To colBoxes.Count - 1
                Dim b = colBoxes(ci)
                Dim gapL = If(ci = 0, 0, b.XMm - (colBoxes(ci - 1).XMm + colBoxes(ci - 1).WidthMm))
                Dim bw = GlobalSettings.BorderWidthPt * MmPerPt
                sb.AppendLine(Fmt("<div class=""col box"" style=""width:{0}mm;margin-left:{1}mm;margin-top:{2}mm;height:{3}mm"">", R2(b.WidthMm), R2(gapL), R2(b.YMm - top), R2(b.HeightMm)))
                Dim inRuns = runs.Where(Function(x) RunInside(x, b)).ToList()
                ' shape milik kolom = titik tengahnya di dalam kotak kolom (bar judul boleh sedikit melewati tepi kolom)
                Dim inShapes = shapes.Where(Function(s) s IsNot b AndAlso ShapeInside(s, b)).ToList()
                RenderBlocks(sb, inRuns, inShapes, b.XMm + bw, b.XMm + b.WidthMm - bw, b.YMm + bw)
                sb.AppendLine("</div>")
            Next
            sb.AppendLine("</div>")

            If below.Count > 0 Then
                RenderBlocks(sb, below, outShapes.Where(Function(s) s.YMm + s.HeightMm / 2 > bottom).ToList(), 0, pageW, bottom)
            End If
            Dim leftover = outside.Count - above.Count - below.Count
            If leftover > 0 Then log($"  ! {leftover} run di antara kolom diabaikan")
        Else
            RenderBlocks(sb, runs, shapes, 0, pageW, 0)
        End If

        sb.Append("</div>")
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Susun semua isi satu kontainer sebagai SATU alur, urut dari atas:
    ''' baris teks, bar judul, tabel, kotak, baris "label : nilai" (.kv) dan blok 2 kolom tanpa kotak (.cols).
    ''' Blok non-teks masuk ke alur sebagai "baris blok" (kiri/atas/tinggi diketahui, html sudah jadi),
    ''' sehingga tabel/bar di dalam item list tetap berada di dalam item itu (ParseFlow menempatkannya
    ''' berdasarkan posisi kirinya, sama seperti baris teks).
    ''' Run dipartisi dulu ke blok (berdasarkan titik tengahnya) sebelum dibentuk jadi baris,
    ''' supaya teks dari kolom/sel yang berbeda tidak tergabung dalam satu baris.
    ''' </summary>
    ''' <returns>Posisi bawah blok terakhir (mm, koordinat region) — untuk margin blok berikutnya.</returns>
    Private Function RenderBlocks(sb As StringBuilder, runs As List(Of TextRun), shapes As List(Of Shape), cLeft As Double, cRight As Double, cTop As Double) As Double
        If runs.Count = 0 Then Return cTop

        ' ----- 1. bar judul, tabel (grid), kotak; sisa run = teks alur -----
        Dim bars As List(Of Shape) = Nothing, tables As List(Of TableGrid) = Nothing, boxes As List(Of Shape) = Nothing
        Dim barRuns As Dictionary(Of Shape, List(Of TextRun)) = Nothing, boxRuns As Dictionary(Of Shape, List(Of TextRun)) = Nothing
        Dim tableRuns As Dictionary(Of TableGrid, List(Of TextRun)) = Nothing
        Dim pool = Partition(runs, shapes, bars, barRuns, tables, tableRuns, boxes, boxRuns)

        ' ----- 0. blok 2 kolom tanpa kotak (mis. label:nilai kiri & kanan) -----
        Dim bands = DetectBands(BuildLines(pool), tables, bars, boxes)
        If bands.Count > 0 Then
            For Each bd In bands
                Dim b = bd
                b.Runs = runs.Where(Function(x) RunCenterY(x) >= b.Top - 0.5 AndAlso RunCenterY(x) <= b.Bottom + 0.5).ToList()
                b.Shapes = shapes.Where(Function(sh) sh.YMm + sh.HeightMm / 2 >= b.Top - 1 AndAlso sh.YMm + sh.HeightMm / 2 <= b.Bottom + 1 AndAlso
                                                     (sh.XMm + sh.WidthMm <= b.Gutter + 0.3 OrElse sh.XMm >= b.Gutter - 0.3)).ToList()
            Next
            Dim bandRuns As New HashSet(Of TextRun)(bands.SelectMany(Function(b) b.Runs))
            Dim bandShapes As New HashSet(Of Shape)(bands.SelectMany(Function(b) b.Shapes))
            runs = runs.Where(Function(x) Not bandRuns.Contains(x)).ToList()
            shapes = shapes.Where(Function(sh) Not bandShapes.Contains(sh)).ToList()
            pool = Partition(runs, shapes, bars, barRuns, tables, tableRuns, boxes, boxRuns)
        End If

        ' ----- 2. konteks teks -----
        Dim flow = BuildLines(pool)
        Dim ctx As New Ctx With {.Top = cTop, .PrevBottom = cTop}
        Dim sizeSrc = If(flow.Any(), flow, BuildLines(runs))
        ctx.Size = DominantSize(sizeSrc)
        ctx.LineH = LinePitch(flow, ctx.Size)
        Dim lefts = flow.Select(Function(l) l.Left).Concat(bands.Select(Function(b) b.Left)).ToList()
        Dim rights = flow.Select(Function(l) l.Right).Concat(bands.Select(Function(b) b.Right)).ToList()
        ctx.Left = If(lefts.Any(), lefts.Min(), cLeft)
        ctx.Right = If(rights.Any(), rights.Max(), cRight)
        For Each l In flow
            l.Top = LineTop(l.Baseline, ctx.Size, ctx.LineH) : l.Height = ctx.LineH
        Next

        ' ----- 3. baris "label : nilai" (nilai boleh berupa tabel) -----
        Dim usedTables As New HashSet(Of TableGrid)
        Dim items As New List(Of Line)
        items.AddRange(DetectKvRows(flow, tables, tableRuns, ctx, usedTables, cRight))
        items.AddRange(flow.Where(Function(l) Not l.Consumed))

        ' ----- 4. blok lain sebagai baris blok -----
        For Each b In bars
            items.Add(BlockLine(BarHtml(b, barRuns(b), ctx), b.XMm, b.YMm, b.WidthMm, b.HeightMm))
        Next
        For Each t In tables.Where(Function(x) Not usedTables.Contains(x))
            items.Add(BlockLine(TableHtml(t, tableRuns(t), ctx.Size), t.BB.XMm, t.BB.YMm, t.BB.WidthMm, t.BB.HeightMm))
        Next
        For Each b In boxes
            items.Add(BlockLine(BoxHtml(b, boxRuns(b)), b.XMm, b.YMm, b.WidthMm, b.HeightMm))
        Next
        For Each bd In bands
            Dim bottom As Double
            Dim html = BandHtml(bd, ctx.Right, bottom)
            items.Add(BlockLine(html, bd.Left, bd.Top, bd.Right - bd.Left, bottom - bd.Top))
        Next
        items = items.OrderBy(Function(l) l.Top).ToList()

        sb.AppendLine(Fmt("<div class=""flow"" style=""padding:0 {0}mm 0 {1}mm;font-size:{2}pt;line-height:{3}mm"">",
                        R2(Math.Max(0, cRight - ctx.Right)), R2(Math.Max(0, ctx.Left - cLeft)), R2(ctx.Size), R2(ctx.LineH)))
        Dim root = ParseFlow(items, ctx)
        EmitNode(sb, root, ctx, 0)
        sb.AppendLine("</div>")
        Return ctx.PrevBottom
    End Function

    ''' <summary>Pisahkan bar judul, tabel (grid dari sel fill + garis), kotak, beserta run di dalamnya. Mengembalikan run sisa (teks alur).</summary>
    Private Function Partition(runs As List(Of TextRun), shapes As List(Of Shape),
                               ByRef bars As List(Of Shape), ByRef barRuns As Dictionary(Of Shape, List(Of TextRun)),
                               ByRef tables As List(Of TableGrid), ByRef tableRuns As Dictionary(Of TableGrid, List(Of TextRun)),
                               ByRef boxes As List(Of Shape), ByRef boxRuns As Dictionary(Of Shape, List(Of TextRun))) As List(Of TextRun)
        bars = shapes.Where(Function(s) s.Kind = ShapeKind.FillRect AndAlso IsTitleBar(s)).ToList()
        Dim cells = shapes.Where(Function(s) s.Kind = ShapeKind.FillRect AndAlso Not IsTitleBar(s) AndAlso Not IsThinLine(s)).ToList()
        ' garis border tabel (fill/stroke tipis) → melengkapi grid: baris judul tanpa background ikut masuk tabel
        Dim gridLines = shapes.Where(Function(s) (s.Kind = ShapeKind.FillRect OrElse s.Kind = ShapeKind.StrokeRect) AndAlso
                                                 Not IsTitleBar(s) AndAlso IsThinLine(s)).ToList()
        tables = ClusterTables(cells).Select(Function(grp) BuildGrid(grp, gridLines)).ToList()
        boxes = shapes.Where(Function(s) s.Kind = ShapeKind.StrokeRect).ToList()

        Dim pool = runs.ToList()
        barRuns = New Dictionary(Of Shape, List(Of TextRun))
        For Each b In bars
            Dim bb = b
            barRuns(b) = pool.Where(Function(x) RunInside(x, bb)).ToList()
            pool = pool.Except(barRuns(b)).ToList()
        Next
        tableRuns = New Dictionary(Of TableGrid, List(Of TextRun))
        For Each t In tables
            Dim tt = t
            tableRuns(t) = pool.Where(Function(x) RunInside(x, tt.BB)).ToList()
            pool = pool.Except(tableRuns(t)).ToList()
        Next
        boxRuns = New Dictionary(Of Shape, List(Of TextRun))
        For Each b In boxes
            Dim bb = b
            boxRuns(b) = pool.Where(Function(x) RunInside(x, bb)).ToList()
            pool = pool.Except(boxRuns(b)).ToList()
        Next
        Return pool
    End Function

    Private Function BlockLine(html As String, left As Double, top As Double, width As Double, height As Double) As Line
        Return New Line With {.BlockHtml = html, .Left = left, .Right = left + width, .Top = top, .Height = height, .Baseline = top}
    End Function

    ''' <summary>Isi margin-top/margin-left (placeholder __MT__/__ML__) pada elemen terluar html blok.</summary>
    Private Function ApplyMargins(html As String, mt As Double, ml As Double) As String
        Dim i = html.IndexOf("__MT__", StringComparison.Ordinal)
        If i >= 0 Then html = html.Substring(0, i) & R2(mt) & html.Substring(i + 6)
        i = html.IndexOf("__ML__", StringComparison.Ordinal)
        If i >= 0 Then html = html.Substring(0, i) & R2(ml) & html.Substring(i + 6)
        Return html
    End Function

    ' ------------------------------------------------------------------
    Private Function BarHtml(b As Shape, rs As List(Of TextRun), ctx As Ctx) As String
        Dim ls = BuildLines(rs)
        Dim txt = String.Join(" ", ls.Select(Function(l) InlineHtml(l, ctx.Size, True)))
        Dim size = If(ls.Any(), ls(0).Size, ctx.Size)
        Return Fmt("<div class=""bar"" style=""margin-top:__MT__mm;margin-left:__ML__mm;width:{0}mm;height:{1}mm;line-height:{1}mm;font-size:{2}pt"">{3}</div>",
                   R2(b.WidthMm), R2(b.HeightMm), R2(size), txt)
    End Function

    Private Function BoxHtml(b As Shape, rs As List(Of TextRun)) As String
        Dim inner As New StringBuilder()
        Dim bw = GlobalSettings.BorderWidthPt * MmPerPt
        RenderBlocks(inner, rs, New List(Of Shape), b.XMm + bw, b.XMm + b.WidthMm - bw, b.YMm + bw)
        Return Fmt("<div class=""box"" style=""margin-top:__MT__mm;margin-left:__ML__mm;width:{0}mm;height:{1}mm"">", R2(b.WidthMm), R2(b.HeightMm)) &
               vbLf & inner.ToString() & "</div>"
    End Function

    ' ------------------------------------------------------------------
    '  BLOK 2 KOLOM TANPA KOTAK
    '  Deteksi: baris dengan celah lebar (≥ 6 mm, bukan celah di sekitar ":" label:nilai) jadi bibit;
    '  diperluas ke bawah/atas selama tidak ada baris/shape yang melintasi celah itu (celah boleh
    '  menyempit sampai 6 mm). Sah kalau ≥ 4 baris dan ≥ 2 baris berisi di kedua sisi.
    '  Render: .cols dengan dua .col (tanpa border), tiap kolom alur sendiri (RenderBlocks rekursif).
    ' ------------------------------------------------------------------
    Private Class Band
        Public Top, Bottom, Left, Right, Gutter As Double
        Public Runs As New List(Of TextRun)
        Public Shapes As New List(Of Shape)
    End Class

    Private Class Elem
        Public Top, Bottom, Left, Right As Double
        Public Line As Line                      ' baris teks; Nothing = shape
        Public Shapes As New List(Of Shape)
        ''' <summary>Rentang x yang ditempati: segmen baris yang berdekatan (celah &lt; 2 mm, kira-kira satu spasi) digabung jadi satu bagian.</summary>
        Public ReadOnly Property Parts As List(Of Tuple(Of Double, Double))
            Get
                Dim res As New List(Of Tuple(Of Double, Double))
                If Line Is Nothing Then
                    res.Add(Tuple.Create(Left, Right))
                    Return res
                End If
                Dim a As Double = Double.NaN, b As Double = Double.NaN
                For Each s In Line.Segs
                    If Double.IsNaN(a) OrElse s.XMm - b >= 2 Then
                        If Not Double.IsNaN(a) Then res.Add(Tuple.Create(a, b))
                        a = s.XMm : b = s.XMm + s.WidthMm
                    Else
                        b = Math.Max(b, s.XMm + s.WidthMm)
                    End If
                Next
                If Not Double.IsNaN(a) Then res.Add(Tuple.Create(a, b))
                Return res
            End Get
        End Property
    End Class

    Private Const BandMinGap As Double = 4.5

    Private Function DetectBands(lines As List(Of Line), tables As List(Of TableGrid), bars As List(Of Shape), boxes As List(Of Shape)) As List(Of Band)
        Dim bands As New List(Of Band)
        If lines.Count < 4 Then Return bands

        Dim elems As New List(Of Elem)
        For Each l In lines
            elems.Add(New Elem With {.Line = l, .Left = l.Left, .Right = l.Right,
                                     .Top = l.Baseline - 0.9 * l.Size * MmPerPt, .Bottom = l.Baseline + 0.25 * l.Size * MmPerPt})
        Next
        For Each t In tables
            elems.Add(New Elem With {.Left = t.BB.XMm, .Right = t.BB.XMm + t.BB.WidthMm, .Top = t.BB.YMm, .Bottom = t.BB.YMm + t.BB.HeightMm})
        Next
        For Each sh In bars.Concat(boxes)
            elems.Add(New Elem With {.Left = sh.XMm, .Right = sh.XMm + sh.WidthMm, .Top = sh.YMm, .Bottom = sh.YMm + sh.HeightMm})
        Next
        elems = elems.OrderBy(Function(e) e.Top).ToList()

        Dim used As New HashSet(Of Elem)
        For idx = 0 To elems.Count - 1
            Dim seed = elems(idx)
            If used.Contains(seed) OrElse seed.Line Is Nothing Then Continue For
            Dim gap = SeedGap(seed.Line)
            If gap Is Nothing Then Continue For
            Dim gL = gap.Item1, gR = gap.Item2
            Dim members As New List(Of Elem) From {seed}
            Dim lastBottom = seed.Bottom
            For k = idx + 1 To elems.Count - 1           ' turun
                Dim e = elems(k)
                If used.Contains(e) OrElse e.Top - lastBottom > 15 OrElse Not Narrow(e, gL, gR) Then
                    If GlobalSettings.DebugBands Then log?.Invoke(Fmt("    bibit «{0}» berhenti di y {1} x {2}–{3} [{4}] celah {5}–{6}: {7}", seed.Line.Text, R2(e.Top), R2(e.Left), R2(e.Right),
                                                                     If(used.Contains(e), "dipakai", If(e.Top - lastBottom > 15, "jarak", "melintas")), R2(gL), R2(gR), If(e.Line IsNot Nothing, e.Line.Text, "[shape]")))
                    Exit For
                End If
                members.Add(e) : lastBottom = Math.Max(lastBottom, e.Bottom)
            Next
            Dim firstTop = seed.Top
            For k = idx - 1 To 0 Step -1                 ' naik
                Dim e = elems(k)
                If used.Contains(e) OrElse firstTop - e.Bottom > 15 OrElse Not Narrow(e, gL, gR) Then Exit For
                members.Add(e) : firstTop = Math.Min(firstTop, e.Top)
            Next

            Dim gutter = (gL + gR) / 2
            Dim lm = members.Where(Function(m) m.Line IsNot Nothing).ToList()
            Dim two = lm.Where(Function(m) m.Left < gutter AndAlso m.Right > gutter).Count()
            Dim leftN = lm.Where(Function(m) m.Right <= gutter).Count(), rightN = lm.Where(Function(m) m.Left >= gutter).Count()
            If lm.Count < 4 OrElse Not (two >= 2 OrElse (two >= 1 AndAlso leftN >= 2 AndAlso rightN >= 2)) Then Continue For
            ' kedua sisi cukup lebar (>= 25% lebar blok) - bukan kolom marker list
            Dim bLeft = members.Min(Function(m) m.Left), bRight = members.Max(Function(m) m.Right)
            If gutter - bLeft < 0.25 * (bRight - bLeft) OrElse bRight - gutter < 0.25 * (bRight - bLeft) Then Continue For
            ' tiap sisi harus berupa kolom teks: mayoritas barisnya (>= 60%) mulai di x yang sama (bukan teks rata tengah / tabel tanpa garis)
            If Not AlignedSide(lm, gutter, True) OrElse Not AlignedSide(lm, gutter, False) Then Continue For

            Dim bd As New Band With {.Gutter = gutter, .Top = members.Min(Function(m) m.Top), .Bottom = members.Max(Function(m) m.Bottom),
                                     .Left = members.Min(Function(m) m.Left), .Right = members.Max(Function(m) m.Right)}
            For Each m In members : used.Add(m) : Next
            bands.Add(bd)
            If GlobalSettings.DebugBands Then
                For Each m In members.OrderBy(Function(x) x.Top)
                    log?.Invoke(Fmt("      y {0} x {1}–{2}: {3}", R2(m.Top), R2(m.Left), R2(m.Right), If(m.Line IsNot Nothing, m.Line.Text, "[shape]")))
                Next
            End If
            log?.Invoke(Fmt("  blok 2 kolom: y {0}–{1} mm, batas x {2} mm ({3} baris, {4} dua sisi) «{5}»", R2(bd.Top), R2(bd.Bottom), R2(gutter), lm.Count, two, seed.Line.Text))
        Next
        Return bands
    End Function

    ''' <summary>Baris di satu sisi celah: >= 60% mulai di x yang sama (+/- 1 mm)?</summary>
    Private Function AlignedSide(lm As List(Of Elem), gutter As Double, leftSide As Boolean) As Boolean
        Dim lefts As New List(Of Double)
        For Each m In lm
            Dim segs = m.Line.Segs.Where(Function(sg) If(leftSide, sg.XMm + sg.WidthMm / 2 < gutter, sg.XMm + sg.WidthMm / 2 >= gutter)).ToList()
            If segs.Count > 0 Then lefts.Add(segs.Min(Function(sg) sg.XMm))
        Next
        If lefts.Count = 0 Then Return True
        Dim best = lefts.Max(Function(x) lefts.Where(Function(y) Math.Abs(y - x) <= 1).Count())
        Return best >= 0.6 * lefts.Count
    End Function

    ''' <summary>Celah terlebar (≥ BandMinGap) antar segmen satu baris, kecuali celah di sekitar ":" (label : nilai).</summary>
    Private Function SeedGap(l As Line) As Tuple(Of Double, Double)
        Dim best As Tuple(Of Double, Double) = Nothing
        For i = 1 To l.Segs.Count - 1
            Dim a = l.Segs(i - 1), b = l.Segs(i)
            If IsColonRun(a) OrElse IsColonRun(b) Then Continue For
            If i = 1 AndAlso IsMarkerText(a.Text) Then Continue For        ' celah marker list -> isi bukan celah kolom
            Dim gap = b.XMm - (a.XMm + a.WidthMm)
            If gap >= BandMinGap AndAlso (best Is Nothing OrElse gap > best.Item2 - best.Item1) Then best = Tuple.Create(a.XMm + a.WidthMm, b.XMm)
        Next
        Return best
    End Function

    Private Function IsColonRun(r As TextRun) As Boolean
        Return r.Words.Count = 1 AndAlso r.Words(0).Text = ":"
    End Function

    ''' <summary>Sempitkan celah [gL,gR] oleh elemen e; False kalau e melintasi celah atau celah jadi &lt; BandMinGap.</summary>
    Private Function Narrow(e As Elem, ByRef gL As Double, ByRef gR As Double) As Boolean
        Dim nL = gL, nR = gR
        For Each p In e.Parts
            Dim a = p.Item1, b = p.Item2
            If b <= nL + 0.3 OrElse a >= nR - 0.3 Then Continue For
            If a <= nL + 0.3 AndAlso b >= nR - 0.3 Then Return False
            If a <= nL + 0.3 Then
                nL = b
            ElseIf b >= nR - 0.3 Then
                nR = a
            ElseIf a - nL < nR - b Then      ' di dalam celah: ikut sisi yang lebih dekat
                nL = b
            Else
                nR = a
            End If
        Next
        If nR - nL < BandMinGap Then Return False
        gL = nL : gR = nR
        Return True
    End Function

    Private Function BandHtml(bd As Band, parentRight As Double, ByRef bottom As Double) As String
        Dim leftRuns = bd.Runs.Where(Function(x) x.XMm + x.WidthMm / 2 < bd.Gutter).ToList()
        Dim rightRuns = bd.Runs.Where(Function(x) x.XMm + x.WidthMm / 2 >= bd.Gutter).ToList()
        Dim leftShapes = bd.Shapes.Where(Function(s) s.XMm + s.WidthMm / 2 < bd.Gutter).ToList()
        Dim rightShapes = bd.Shapes.Where(Function(s) s.XMm + s.WidthMm / 2 >= bd.Gutter).ToList()
        Dim sbL As New StringBuilder(), sbR As New StringBuilder()
        Dim bL = RenderBlocks(sbL, leftRuns, leftShapes, bd.Left, bd.Gutter, bd.Top)
        Dim bR = RenderBlocks(sbR, rightRuns, rightShapes, bd.Gutter, parentRight, bd.Top)
        bottom = Math.Max(bd.Top, Math.Max(bL, bR))
        Return "<div class=""cols"" style=""margin-top:__MT__mm;margin-left:__ML__mm"">" & vbLf &
               Fmt("<div class=""col"" style=""width:{0}mm"">", R2(bd.Gutter - bd.Left)) & vbLf & sbL.ToString() & "</div>" & vbLf &
               Fmt("<div class=""col"" style=""width:{0}mm"">", R2(parentRight - bd.Gutter)) & vbLf & sbR.ToString() & "</div>" & vbLf &
               "</div>"
    End Function

    ' ------------------------------------------------------------------
    '  BARIS "LABEL : NILAI"  (.kv)
    '  Baris dengan kata ":" berdiri sendiri (bukan kata pertama). Baris berikutnya yang mulai di x label
    '  = lanjutan label, di x nilai = lanjutan nilai; tabel di kanan ":" = nilai. Berhenti di baris
    '  label:nilai berikutnya (":" di x yang sama) atau baris lain.
    '  Render: flex → [label lebar (xColon−xLabel)] [":" lebar (xNilai−xColon)] [nilai sisa lebar].
    ' ------------------------------------------------------------------
    Private Function DetectKvRows(flow As List(Of Line), tables As List(Of TableGrid), tableRuns As Dictionary(Of TableGrid, List(Of TextRun)),
                                  ctx As Ctx, usedTables As HashSet(Of TableGrid), cRight As Double) As List(Of Line)
        Dim result As New List(Of Line)
        Dim ls = flow.OrderBy(Function(l) l.Baseline).ToList()
        Dim spaceMm = GlobalSettings.SpaceWidthEm * ctx.Size * MmPerPt

        For i = 0 To ls.Count - 1
            Dim L = ls(i)
            If L.Consumed Then Continue For
            Dim keySegs As List(Of TextRun) = Nothing, colon As TextRun = Nothing, valSegs As List(Of TextRun) = Nothing
            If Not SplitStandaloneColon(L, keySegs, colon, valSegs) Then Continue For
            Dim labelX = L.Left, cX = colon.XMm
            If cX - labelX < 5 OrElse cX - labelX > 70 Then Continue For
            Dim vX As Double = If(valSegs.Any(), valSegs(0).XMm, Double.NaN)

            Dim labelLines As New List(Of Line) From {LineOf(keySegs, L)}
            Dim valueItems As New List(Of Object)        ' Line (teks nilai) atau TableGrid, urut dari atas
            If valSegs.Any() Then valueItems.Add(LineOf(valSegs, L))
            Dim rowBottom = L.Top + L.Height
            Dim usedLines As New List(Of Line)

            ' kandidat lanjutan: baris & tabel setelah baris ini, urut posisi atas
            Dim cand As New List(Of Tuple(Of Double, Object))
            For j = i + 1 To ls.Count - 1
                If Not ls(j).Consumed Then cand.Add(Tuple.Create(ls(j).Top, CObj(ls(j))))
            Next
            For Each t In tables
                If Not usedTables.Contains(t) AndAlso t.BB.YMm >= L.Top - 1 AndAlso t.BB.XMm > cX Then cand.Add(Tuple.Create(t.BB.YMm, CObj(t)))
            Next
            For Each c In cand.OrderBy(Function(x) x.Item1)
                If c.Item1 - rowBottom > ctx.LineH * 1.5 Then Exit For
                If TypeOf c.Item2 Is Line Then
                    Dim n = DirectCast(c.Item2, Line)
                    Dim k2 As List(Of TextRun) = Nothing, c2 As TextRun = Nothing, v2 As List(Of TextRun) = Nothing
                    If SplitStandaloneColon(n, k2, c2, v2) AndAlso Math.Abs(c2.XMm - cX) < 3 Then Exit For
                    ' segmen di kiri ":" = lanjutan label, di x nilai = lanjutan nilai (satu baris bisa berisi keduanya)
                    Dim labSegs = n.Segs.Where(Function(sg) sg.XMm + sg.WidthMm <= cX + 0.5).ToList()
                    Dim valSegs2 = n.Segs.Where(Function(sg) sg.XMm > cX + 0.5).ToList()
                    Dim okLab = labSegs.Count > 0 AndAlso Math.Abs(labSegs(0).XMm - labelX) <= 0.7
                    Dim okVal = valSegs2.Count > 0 AndAlso Not Double.IsNaN(vX) AndAlso Math.Abs(valSegs2(0).XMm - vX) <= 0.7
                    If Not okLab AndAlso Not okVal Then Exit For
                    If (labSegs.Count > 0 AndAlso Not okLab) OrElse (valSegs2.Count > 0 AndAlso Not okVal) Then Exit For
                    If okLab Then labelLines.Add(LineOf(labSegs, n))
                    If okVal Then valueItems.Add(LineOf(valSegs2, n))
                    usedLines.Add(n)
                    rowBottom = Math.Max(rowBottom, n.Top + n.Height)
                Else
                    Dim t = DirectCast(c.Item2, TableGrid)
                    If Double.IsNaN(vX) Then vX = t.BB.XMm
                    If Math.Abs(t.BB.XMm - vX) > 3 Then Exit For
                    valueItems.Add(t) : usedTables.Add(t)
                    rowBottom = Math.Max(rowBottom, t.BB.YMm + t.BB.HeightMm)
                End If
            Next
            If Double.IsNaN(vX) Then
                For Each t In valueItems.OfType(Of TableGrid)() : usedTables.Remove(t) : Next
                Continue For
            End If

            ' ---- html ----
            Dim sb As New StringBuilder()
            sb.AppendLine("<div class=""kv"" style=""display:flex;align-items:flex-start;margin-top:__MT__mm;margin-left:__ML__mm"">")
            sb.AppendLine(Fmt("<div style=""flex:0 0 {0}mm"">", R2(cX - labelX)))
            sb.Append(ParasHtml(labelLines, ctx, L.Top, cX, spaceMm))
            sb.AppendLine("</div>")
            sb.AppendLine(Fmt("<div style=""flex:0 0 {0}mm"">:</div>", R2(vX - cX)))
            sb.AppendLine("<div style=""flex:1;min-width:0"">")
            Dim prevB = L.Top
            Dim pend As New List(Of Line)
            For Each o In valueItems
                If TypeOf o Is Line Then
                    pend.Add(DirectCast(o, Line))
                Else
                    If pend.Count > 0 Then sb.Append(ParasHtml(pend, ctx, prevB, cRight, spaceMm)) : prevB = pend.Last().Top + pend.Last().Height : pend.Clear()
                    Dim t = DirectCast(o, TableGrid)
                    sb.AppendLine(ApplyMargins(TableHtml(t, tableRuns(t), ctx.Size), Math.Max(0, t.BB.YMm - prevB), Math.Max(0, t.BB.XMm - vX)))
                    prevB = t.BB.YMm + t.BB.HeightMm
                End If
            Next
            If pend.Count > 0 Then sb.Append(ParasHtml(pend, ctx, prevB, cRight, spaceMm))   ' tepi kanan kontainer: nilai rata kiri, pemisahan paragraf tak mengubah tampilan
            sb.AppendLine("</div>")
            sb.Append("</div>")

            L.Consumed = True
            For Each u In usedLines : u.Consumed = True : Next
            result.Add(New Line With {.BlockHtml = sb.ToString(), .Left = labelX, .Right = ctx.Right, .Top = L.Top, .Height = rowBottom - L.Top, .Baseline = L.Top})
            log?.Invoke(Fmt("  label:nilai «{0}» ({1} baris label, {2} item nilai)", labelLines(0).Text, labelLines.Count, valueItems.Count))
        Next
        Return result
    End Function

    ''' <summary>Kata ":" berdiri sendiri (bukan kata pertama) → key / colon / value.</summary>
    Private Function SplitStandaloneColon(l As Line, ByRef keySegs As List(Of TextRun), ByRef colon As TextRun, ByRef valSegs As List(Of TextRun)) As Boolean
        keySegs = New List(Of TextRun) : valSegs = New List(Of TextRun) : colon = Nothing
        For Each sg In l.Segs
            If sg.IsSuper Then Continue For
            If colon IsNot Nothing Then valSegs.Add(sg) : Continue For
            Dim ci = sg.Words.FindIndex(Function(w) w.Text = ":")
            If ci < 0 Then keySegs.Add(sg) : Continue For
            If ci > 0 Then keySegs.Add(SubRun(sg, 0, ci))
            colon = SubRun(sg, ci, ci + 1)
            If ci + 1 < sg.Words.Count Then valSegs.Add(SubRun(sg, ci + 1, sg.Words.Count))
        Next
        Return colon IsNot Nothing AndAlso keySegs.Count > 0
    End Function

    Private Function LineOf(segs As List(Of TextRun), src As Line) As Line
        Dim l As New Line With {.Baseline = src.Baseline, .Top = src.Top, .Height = src.Height,
                                .Left = segs(0).XMm, .Right = segs.Last().XMm + segs.Last().WidthMm, .Size = segs.Max(Function(s) s.SizePt)}
        l.Segs.AddRange(segs)
        Return l
    End Function

    ''' <summary>Baris-baris → paragraf rata kiri (baris digabung kalau kata pertamanya tidak muat di baris sebelumnya).</summary>
    Private Function ParasHtml(ls As List(Of Line), ctx As Ctx, prevBottom As Double, rightEdge As Double, spaceMm As Double) As String
        Dim sb As New StringBuilder()
        Dim para As New List(Of Line)
        Dim prev As Line = Nothing
        Dim paraTop As Double = 0
        Dim flush = Sub()
                        If para.Count = 0 Then Return
                        Dim allBold = para.All(Function(x) x.AllBold)
                        Dim st As New StringBuilder("text-align:left;")
                        If paraTop - prevBottom > 0.05 Then st.Append(Fmt("margin-top:{0}mm;", R2(paraTop - prevBottom)))
                        If Math.Abs(para(0).Size - ctx.Size) > 0.2 Then st.Append(Fmt("font-size:{0}pt;", R2(para(0).Size)))
                        sb.AppendLine(Fmt("<p{0} style=""{1}"">{2}</p>", If(allBold, " class=""sub""", ""), st.ToString(),
                                          String.Join(" ", para.Select(Function(x) InlineHtml(x, ctx.Size, allBold)))))
                        prevBottom = para.Last().Top + para.Last().Height
                        para.Clear()
                    End Sub
        For Each l In ls.OrderBy(Function(x) x.Baseline)
            Dim cont = prev IsNot Nothing AndAlso para.Count > 0 AndAlso l.Top - (prev.Top + prev.Height) < ctx.LineH * 0.6 AndAlso
                       Not FitsOnPrev(prev, l, rightEdge, spaceMm) AndAlso Not (prev.AllBold Xor l.AllBold)
            If Not cont Then flush() : paraTop = l.Top
            para.Add(l)
            prev = l
        Next
        flush()
        Return sb.ToString()
    End Function

    ' ------------------------------------------------------------------
    '  TABEL: sel isi = rect fill (dikelompokkan yang bersentuhan); grid lengkap dibangun BuildGrid
    ' ------------------------------------------------------------------
    Private Function ClusterTables(cells As List(Of Shape)) As List(Of List(Of Shape))
        Dim result As New List(Of List(Of Shape))
        Dim remaining = cells.ToList()
        While remaining.Count > 0
            Dim grp As New List(Of Shape) From {remaining(0)}
            remaining.RemoveAt(0)
            Dim changed = True
            While changed
                changed = False
                For i = remaining.Count - 1 To 0 Step -1
                    Dim c = remaining(i)
                    If grp.Any(Function(g) Touch(g, c)) Then
                        grp.Add(c) : remaining.RemoveAt(i) : changed = True
                    End If
                Next
            End While
            If grp.Count >= 2 Then result.Add(grp)
        End While
        Return result
    End Function

    ' ------------------------------------------------------------------
    '  GRID TABEL: sel isi (fill) + garis border tipis → batas baris/kolom.
    '  Baris judul biasanya TIDAK punya background (hanya border), jadi tabel
    '  diperluas ke atas/bawah mengikuti garis vertikal yang menyambung ke sel isi.
    '  Area tanpa fill di dalam grid digabung jadi sel (colspan/rowspan) selama
    '  tidak dipisahkan garis.
    ' ------------------------------------------------------------------
    Private Class GridCell
        Public R0, R1, C0, C1 As Integer     ' baris [R0,R1), kolom [C0,C1)
        Public Fill As Shape                 ' Nothing = sel tanpa background (mis. baris judul)
    End Class

    Private Class TableGrid
        Public Rows As New List(Of Double)   ' batas baris (mm), Count = jumlah baris + 1
        Public Cols As New List(Of Double)   ' batas kolom (mm), Count = jumlah kolom + 1
        Public Cells As New List(Of GridCell)
        Public BB As Shape
        Public FirstFilledRow As Integer     ' baris pertama yang punya sel fill (baris di atasnya = judul)
    End Class

    Private Const GridTol As Double = 0.6

    ''' <summary>Rect setipis garis (≤ ThinLineMaxPt) dan cukup panjang → garis border tabel.</summary>
    Private Function IsThinLine(s As Shape) As Boolean
        Return Math.Min(s.WidthMm, s.HeightMm) * 72 / 25.4 <= GlobalSettings.ThinLineMaxPt AndAlso
               Math.Max(s.WidthMm, s.HeightMm) > 1.5
    End Function

    Private Function BuildGrid(cells As List(Of Shape), lines As List(Of Shape)) As TableGrid
        Dim bb = BBox(cells)
        Dim x0 = bb.XMm, x1 = bb.XMm + bb.WidthMm, y0 = bb.YMm, y1 = bb.YMm + bb.HeightMm
        Dim vert = lines.Where(Function(l) l.HeightMm > l.WidthMm).ToList()
        Dim horz = lines.Where(Function(l) l.WidthMm >= l.HeightMm).ToList()

        ' perluas ke atas/bawah mengikuti garis vertikal yang menyambung ke tabel (lebar tetap)
        For pass = 1 To 10
            Dim ny0 = y0, ny1 = y1
            For Each l In vert
                If l.XMm < x0 - GridTol OrElse l.XMm > x1 + GridTol Then Continue For
                If l.YMm > y1 + GridTol OrElse l.YMm + l.HeightMm < y0 - GridTol Then Continue For
                ' garis yang cuma melewati tepi setebal garisnya (± 1 mm) bukan baris baru
                If l.YMm < y0 - 1 Then ny0 = Math.Min(ny0, l.YMm)
                If l.YMm + l.HeightMm > y1 + 1 Then ny1 = Math.Max(ny1, l.YMm + l.HeightMm)
            Next
            If Math.Abs(ny0 - y0) < 0.01 AndAlso Math.Abs(ny1 - y1) < 0.01 Then Exit For
            y0 = ny0 : y1 = ny1
        Next

        Dim g As New TableGrid With {.BB = New Shape With {.XMm = x0, .YMm = y0, .WidthMm = x1 - x0, .HeightMm = y1 - y0}}
        Dim vIn = vert.Where(Function(l) l.XMm >= x0 - GridTol AndAlso l.XMm <= x1 + GridTol AndAlso
                                         l.YMm + l.HeightMm >= y0 - GridTol AndAlso l.YMm <= y1 + GridTol).ToList()
        Dim hIn = horz.Where(Function(l) l.YMm >= y0 - GridTol AndAlso l.YMm <= y1 + GridTol AndAlso
                                         l.XMm + l.WidthMm >= x0 - GridTol AndAlso l.XMm <= x1 + GridTol).ToList()

        ' batas baris/kolom = tepi tabel + tepi sel fill; posisi garis hanya menambah batas yang belum ada
        ' (tepi sel lebih akurat: garis digambar dari tepi sel setebal 0.26 mm)
        Dim ys As New List(Of Double) From {y0, y1}
        Dim xs As New List(Of Double) From {x0, x1}
        For Each c In cells
            ys.Add(c.YMm) : ys.Add(c.YMm + c.HeightMm) : xs.Add(c.XMm) : xs.Add(c.XMm + c.WidthMm)
        Next
        ys = Cluster(ys, GridTol) : xs = Cluster(xs, GridTol)
        For Each l In hIn
            Dim v = l.YMm
            If Not ys.Any(Function(e) Math.Abs(e - v) <= GridTol) Then ys.Add(v)
        Next
        For Each l In vIn
            Dim v = l.XMm
            If Not xs.Any(Function(e) Math.Abs(e - v) <= GridTol) Then xs.Add(v)
        Next
        g.Rows = Cluster(ys, GridTol) : g.Cols = Cluster(xs, GridTol)
        Dim nR = g.Rows.Count - 1, nC = g.Cols.Count - 1
        If nR < 1 OrElse nC < 1 Then Return g

        Dim taken(nR - 1, nC - 1) As Boolean
        Dim free = Function(r0 As Integer, r1 As Integer, c0 As Integer, c1 As Integer) As Boolean
                       For r = r0 To r1 - 1
                           For c = c0 To c1 - 1
                               If taken(r, c) Then Return False
                           Next
                       Next
                       Return True
                   End Function
        Dim mark = Sub(gc As GridCell)
                       For r = gc.R0 To gc.R1 - 1
                           For c = gc.C0 To gc.C1 - 1 : taken(r, c) = True : Next
                       Next
                       g.Cells.Add(gc)
                   End Sub

        ' 1) sel fill → unit grid yang ditutupinya
        For Each c In cells.OrderBy(Function(s) s.YMm).ThenBy(Function(s) s.XMm)
            Dim gc As New GridCell With {.Fill = c,
                .R0 = Nearest(g.Rows, c.YMm), .R1 = Nearest(g.Rows, c.YMm + c.HeightMm),
                .C0 = Nearest(g.Cols, c.XMm), .C1 = Nearest(g.Cols, c.XMm + c.WidthMm)}
            gc.R0 = Math.Min(gc.R0, nR - 1) : gc.C0 = Math.Min(gc.C0, nC - 1)
            gc.R1 = Math.Max(gc.R1, gc.R0 + 1) : gc.C1 = Math.Max(gc.C1, gc.C0 + 1)
            If Not free(gc.R0, gc.R1, gc.C0, gc.C1) Then Continue For
            mark(gc)
        Next
        g.FirstFilledRow = If(g.Cells.Any(), g.Cells.Min(Function(c) c.R0), nR)

        ' 2) sisa unit (tanpa fill) → gabung ke kanan/bawah selama tidak ada garis pemisah
        For r = 0 To nR - 1
            For c = 0 To nC - 1
                If taken(r, c) Then Continue For
                Dim c1 = c + 1
                While c1 < nC AndAlso Not taken(r, c1) AndAlso Not HasVLine(vIn, g, c1, r, r + 1)
                    c1 += 1
                End While
                Dim r1 = r + 1
                While r1 < nR AndAlso free(r1, r1 + 1, c, c1) AndAlso Not HasHLine(hIn, g, r1, c, c1)
                    r1 += 1
                End While
                mark(New GridCell With {.R0 = r, .R1 = r1, .C0 = c, .C1 = c1})
            Next
        Next
        g.Cells = g.Cells.OrderBy(Function(x) x.R0).ThenBy(Function(x) x.C0).ToList()
        Return g
    End Function

    ''' <summary>Ada garis vertikal di batas kolom ci yang menutup baris [rFrom, rTo)?</summary>
    Private Function HasVLine(vIn As List(Of Shape), g As TableGrid, ci As Integer, rFrom As Integer, rTo As Integer) As Boolean
        Dim x = g.Cols(ci)
        Dim segs = vIn.Where(Function(l) Math.Abs(l.XMm - x) <= GridTol).Select(Function(l) Tuple.Create(l.YMm, l.YMm + l.HeightMm))
        Return Covered(segs, g.Rows(rFrom), g.Rows(rTo))
    End Function

    ''' <summary>Ada garis horizontal di batas baris ri yang menutup kolom [cFrom, cTo)?</summary>
    Private Function HasHLine(hIn As List(Of Shape), g As TableGrid, ri As Integer, cFrom As Integer, cTo As Integer) As Boolean
        Dim y = g.Rows(ri)
        Dim segs = hIn.Where(Function(l) Math.Abs(l.YMm - y) <= GridTol).Select(Function(l) Tuple.Create(l.XMm, l.XMm + l.WidthMm))
        Return Covered(segs, g.Cols(cFrom), g.Cols(cTo))
    End Function

    ''' <summary>Gabungan segmen (awal, akhir) menutup rentang [a, b] tanpa celah > toleransi?</summary>
    Private Function Covered(segs As IEnumerable(Of Tuple(Of Double, Double)), a As Double, b As Double) As Boolean
        Dim cur = a
        For Each s In segs.OrderBy(Function(t) t.Item1)
            If s.Item1 > cur + GridTol Then Exit For
            cur = Math.Max(cur, s.Item2)
            If cur >= b - GridTol Then Return True
        Next
        Return cur >= b - GridTol
    End Function

    ''' <summary>Indeks nilai batas yang paling dekat dengan v.</summary>
    Private Function Nearest(bounds As List(Of Double), v As Double) As Integer
        Dim best = 0
        For i = 1 To bounds.Count - 1
            If Math.Abs(bounds(i) - v) < Math.Abs(bounds(best) - v) Then best = i
        Next
        Return best
    End Function

    ''' <summary>HTML tabel (margin-top/left = placeholder __MT__/__ML__, diisi saat ditempatkan di alur).</summary>
    Private Function TableHtml(g As TableGrid, rs As List(Of TextRun), baseSize As Double) As String
        Dim bb = g.BB
        Dim sb As New StringBuilder()
        sb.AppendLine(Fmt("<table class=""tbl"" style=""margin-top:__MT__mm;margin-left:__ML__mm;width:{0}mm"">", R2(bb.WidthMm)))
        sb.Append("<colgroup>")
        For ci = 0 To g.Cols.Count - 2
            sb.Append(Fmt("<col style=""width:{0}mm"">", R2(g.Cols(ci + 1) - g.Cols(ci))))
        Next
        sb.AppendLine("</colgroup>")

        ' isi & perataan tiap sel dihitung dulu; sel yang perataannya ambigu (teks memenuhi sel) ikut mayoritas kolomnya
        Dim cellLines As New Dictionary(Of GridCell, List(Of Line))
        Dim alignOf As New Dictionary(Of GridCell, String)
        Dim cellHead As New Dictionary(Of GridCell, Boolean)
        For Each cell In g.Cells
            Dim rect = CellRect(g, cell)
            Dim cl = BuildLines(rs.Where(Function(x) RunInside(x, rect)).ToList())
            cellLines(cell) = cl
            ' judul = sel tanpa background di atas baris isi pertama, atau baris pertama yang seluruhnya bold
            cellHead(cell) = (cell.Fill Is Nothing AndAlso cell.R0 < g.FirstFilledRow) OrElse
                             (cell.R0 = 0 AndAlso cl.Any() AndAlso cl.All(Function(l) l.AllBold))
            alignOf(cell) = CellAlign(cl, rect)
        Next
        For Each cell In g.Cells.Where(Function(c) alignOf(c) Is Nothing).ToList()
            Dim c0 = cell.C0, head = cellHead(cell)
            Dim votes = g.Cells.Where(Function(o) o.C0 = c0 AndAlso cellHead(o) = head AndAlso alignOf(o) IsNot Nothing).
                                GroupBy(Function(o) alignOf(o)).OrderByDescending(Function(grp) grp.Count()).FirstOrDefault()
            alignOf(cell) = If(votes Is Nothing, "center", votes.Key)
        Next

        For ri = 0 To g.Rows.Count - 2
            sb.Append(Fmt("<tr style=""height:{0}mm"">", R2(g.Rows(ri + 1) - g.Rows(ri))))
            Dim rowIdx = ri
            For Each cell In g.Cells.Where(Function(x) x.R0 = rowIdx)
                Dim cl = cellLines(cell)
                Dim isHead = cellHead(cell)
                Dim tag = If(isHead, "th", "td")
                Dim attrs As New StringBuilder()
                If cell.C1 - cell.C0 > 1 Then attrs.Append(Fmt(" colspan=""{0}""", cell.C1 - cell.C0))
                If cell.R1 - cell.R0 > 1 Then attrs.Append(Fmt(" rowspan=""{0}""", cell.R1 - cell.R0))
                Dim st As New StringBuilder()
                If cell.Fill IsNot Nothing AndAlso cell.Fill.FillHex IsNot Nothing AndAlso Luminance(cell.Fill.FillHex) < 0.99 Then
                    st.Append("background:" & cell.Fill.FillHex & ";")
                End If
                If alignOf(cell) <> "center" Then st.Append("text-align:" & alignOf(cell) & ";")
                If cl.Any() Then
                    Dim cs = cl(0).Size
                    If Math.Abs(cs - baseSize) > 0.2 Then st.Append(Fmt("font-size:{0}pt;", R2(cs)))
                    st.Append(Fmt("line-height:{0}mm;", R2(LinePitch(cl, cs))))
                End If
                Dim content = String.Join("<br>", cl.Select(Function(l) InlineHtml(l, baseSize, isHead)))
                sb.Append(Fmt("<{0}{1} style=""{2}"">{3}</{0}>", tag, attrs.ToString(), st.ToString(), content))
            Next
            sb.AppendLine("</tr>")
        Next
        sb.Append("</table>")
        Return sb.ToString()
    End Function

    Private Function CellRect(g As TableGrid, cell As GridCell) As Shape
        Return New Shape With {.XMm = g.Cols(cell.C0), .YMm = g.Rows(cell.R0),
                               .WidthMm = g.Cols(cell.C1) - g.Cols(cell.C0), .HeightMm = g.Rows(cell.R1) - g.Rows(cell.R0)}
    End Function

    ''' <summary>
    ''' Perataan teks sel dari posisi baris terakhirnya: "center" / "left" / "right";
    ''' Nothing = ambigu (teks memenuhi lebar sel) atau sel kosong.
    ''' </summary>
    Private Function CellAlign(cl As List(Of Line), rect As Shape) As String
        If cl.Count = 0 Then Return Nothing
        Dim l = cl.Last()
        Dim padL = l.Left - rect.XMm, padR = rect.XMm + rect.WidthMm - l.Right
        If padL < 2.2 AndAlso padR < 2.2 Then Return Nothing
        If Math.Abs(padL - padR) <= 1.2 Then Return "center"
        If padL < 2.2 Then Return "left"
        If padR < 2.2 Then Return "right"
        Return Nothing
    End Function

    ' ------------------------------------------------------------------
    '  LIST / PARAGRAF / BLOK  (alur satu kontainer)
    '  - baris marker (A. / 1. / a. / (1) / 1) / •) + indentasi → list bertingkat
    '  - baris teks → paragraf (digabung selama lanjutan wajar); baris yang menjorok → margin-left /
    '    text-indent; baris "tersebar" (celah antar segmen ≥ 3 mm, tidak sampai tepi kanan) → tiap
    '    segmen diposisikan absolut supaya jaraknya persis (mis. label Rendah/Sedang/Tinggi, tabel tanpa garis)
    '  - baris blok (tabel/bar/kotak/kv/cols) → ditempatkan di item list yang sesuai posisi kirinya
    ' ------------------------------------------------------------------
    Private Function ParseFlow(ls As List(Of Line), ctx As Ctx) As Node
        Dim root As New Node With {.Kind = "root"}
        Dim stack As New List(Of Node)       ' item yang sedang terbuka
        Dim curPara As Node = Nothing
        Dim prev As Line = Nothing
        Dim spaceMm = GlobalSettings.SpaceWidthEm * ctx.Size * MmPerPt

        For Each l In ls
            Dim lineTop = l.Top
            Dim gapTop = lineTop - ctx.PrevBottom          ' jarak dari blok sebelumnya (mm)

            ' ---- baris blok (tabel / bar / kotak / kv / cols) ----
            If l.IsBlock Then
                While stack.Count > 0 AndAlso l.Left < stack.Last().ContentX - 0.7
                    stack.RemoveAt(stack.Count - 1)
                End While
                Dim container = If(stack.Count > 0, stack.Last(), root)
                Dim baseX = If(stack.Count > 0, container.ContentX, ctx.Left)
                container.Add(New Node With {.Kind = "block", .Html = l.BlockHtml, .MarginTop = Math.Max(0, gapTop), .MarginLeft = l.Left - baseX})
                curPara = Nothing : prev = Nothing
                ctx.PrevBottom = l.Top + l.Height
                Continue For
            End If

            Dim first = l.Segs(0)
            Dim isMarker = Not first.IsSuper AndAlso IsMarkerText(first.Text) AndAlso
                           (l.Segs.Count = 1 OrElse l.Segs(1).XMm - (first.XMm + first.WidthMm) > 0.8)

            If isMarker Then
                Dim mX = first.XMm
                Dim cX = If(l.Segs.Count > 1, l.Segs(1).XMm, first.XMm + first.WidthMm + spaceMm * 2)

                While stack.Count > 0 AndAlso stack.Last().MarkerX > mX + 0.7
                    stack.RemoveAt(stack.Count - 1)
                End While

                Dim list As Node
                If stack.Count = 0 OrElse mX > stack.Last().MarkerX + 0.7 Then
                    Dim parent = If(stack.Count = 0, root, stack.Last())
                    list = New Node With {.Kind = "list"}
                    parent.Add(list)
                Else
                    list = stack.Last().Parent
                    stack.RemoveAt(stack.Count - 1)
                End If

                Dim item As New Node With {.Kind = "item", .MarkerX = mX, .ContentX = cX, .MarkerSeg = first, .MarginTop = Math.Max(0, gapTop)}
                list.Add(item)
                stack.Add(item)

                curPara = New Node With {.Kind = "para", .BaseX = cX}
                Dim rest As New Line With {.Baseline = l.Baseline, .Size = l.Size, .Left = cX, .Right = l.Right, .Top = l.Top, .Height = l.Height}
                rest.Segs.AddRange(l.Segs.Skip(1))
                If rest.Segs.Count > 0 Then
                    curPara.IsSpread = IsSpread(rest, ctx)
                    curPara.Lines.Add(rest)
                End If
                item.Add(curPara)
            Else
                ' pop item yang lebih dalam dari posisi baris ini; baris yang menjorok lebih dalam tetap milik item terdalam
                While stack.Count > 0 AndAlso l.Left < stack.Last().ContentX - 0.7
                    stack.RemoveAt(stack.Count - 1) : curPara = Nothing
                End While
                Dim container = If(stack.Count > 0, stack.Last(), root)
                Dim baseX = If(stack.Count > 0, container.ContentX, ctx.Left)
                Dim indent = Math.Max(0, l.Left - baseX)
                Dim spread = IsSpread(l, ctx) OrElse indent > 0.7      ' diposisikan absolut (jarak/indentasi persis)

                ' lanjutan paragraf? syarat: paragraf sama, jarak normal, kata pertama TIDAK muat di baris sebelumnya,
                ' tidak ganti bold, bukan baris menjorok/tersebar, bukan baris penanda <<if …>> / <<Page Break>>
                Dim cont = curPara IsNot Nothing AndAlso curPara.Parent Is container AndAlso curPara.Lines.Count > 0 AndAlso
                           Not curPara.IsSpread AndAlso Not spread AndAlso indent <= 0.7 AndAlso
                           gapTop < ctx.LineH * 0.6 AndAlso
                           prev IsNot Nothing AndAlso Not FitsOnPrev(prev, l, ctx.Right, spaceMm) AndAlso
                           Not (prev.AllBold Xor l.AllBold) AndAlso
                           Not DirectiveProcessor.IsDirective(l.Text) AndAlso Not DirectiveProcessor.IsDirective(prev.Text)
                ' baris tersebar/menjorok berurutan (jarak < 1.5 baris) digabung ke satu blok posisi absolut
                Dim joinSpread = spread AndAlso curPara IsNot Nothing AndAlso curPara.IsSpread AndAlso curPara.Parent Is container AndAlso
                                 prev IsNot Nothing AndAlso lineTop - prev.Top < ctx.LineH * 1.5
                If cont OrElse joinSpread Then
                    curPara.Lines.Add(l)
                Else
                    curPara = New Node With {.Kind = "para", .MarginTop = Math.Max(0, gapTop), .Indent = indent, .IsSpread = spread, .BaseX = baseX}
                    curPara.Lines.Add(l)
                    container.Add(curPara)
                End If
            End If

            ctx.PrevBottom = lineTop + ctx.LineH
            prev = l
        Next
        Return root
    End Function

    ''' <summary>Baris "tersebar": ada celah ≥ 3 mm antar segmen dan baris tidak mencapai tepi kanan (bukan teks justify).</summary>
    Private Function IsSpread(l As Line, ctx As Ctx) As Boolean
        If l.Segs.Count < 2 OrElse l.Right > ctx.Right - 3 Then Return False
        For i = 1 To l.Segs.Count - 1
            If l.Segs(i).XMm - (l.Segs(i - 1).XMm + l.Segs(i - 1).WidthMm) >= 3 Then Return True
        Next
        Return False
    End Function

    ' kata pertama baris ini seharusnya muat di baris sebelumnya → berarti paragraf baru
    Private Function FitsOnPrev(prev As Line, cur As Line, right As Double, spaceMm As Double) As Boolean
        Dim firstWord = cur.Segs(0).Words(0)
        Return prev.Right + spaceMm + firstWord.WidthMm <= right + 0.3
    End Function

    Private Sub EmitNode(sb As StringBuilder, n As Node, ctx As Ctx, depth As Integer)
        Dim ind = New String(" "c, depth * 2)
        Select Case n.Kind
            Case "root"
                For Each c In n.Children : EmitNode(sb, c, ctx, depth) : Next
            Case "list"
                sb.AppendLine(ind & "<ol class=""lst"">")
                For Each c In n.Children : EmitNode(sb, c, ctx, depth + 1) : Next
                sb.AppendLine(ind & "</ol>")
            Case "item"
                Dim mw = n.ContentX - n.MarkerX
                sb.AppendLine(ind & Fmt("<li{0}><span class=""m"" style=""width:{1}mm{2}"">{3}</span><div class=""b"">",
                                      MarginAttr(n.MarginTop), R2(mw),
                                      If(n.MarkerSeg.IsBold, ";font-weight:bold", ""),
                                      WebUtility.HtmlEncode(n.MarkerSeg.Text)))
                For Each c In n.Children : EmitNode(sb, c, ctx, depth + 1) : Next
                sb.AppendLine(ind & "</div></li>")
            Case "block"
                sb.AppendLine(ind & ApplyMargins(n.Html, n.MarginTop, n.MarginLeft))
            Case "para"
                If n.Lines.Count = 0 Then Return
                Dim allBold = n.Lines.All(Function(l) l.AllBold)
                Dim size = n.Lines(0).Size
                Dim st As New StringBuilder()
                If n.MarginTop > 0.05 Then st.Append(Fmt("margin-top:{0}mm;", R2(n.MarginTop)))
                If Math.Abs(size - ctx.Size) > 0.2 Then st.Append(Fmt("font-size:{0}pt;", R2(size)))
                If n.IsSpread Then
                    ' tiap segmen diposisikan absolut (relatif ke kiri-atas blok) — posisi x/y persis seperti PDF
                    Dim top0 = n.Lines(0).Top
                    st.Append(Fmt("position:relative;height:{0}mm;text-align:left;", R2(n.Lines.Last().Top - top0 + ctx.LineH)))
                    Dim parts As New StringBuilder()
                    For Each ln In n.Lines
                        For Each s In ln.Segs
                            parts.Append(Fmt("<span style=""position:absolute;left:{0}mm;top:{1}mm;white-space:nowrap"">{2}</span>",
                                             R2(s.XMm - n.BaseX), R2(ln.Top - top0), InlineSegs(New List(Of TextRun) From {s}, ctx.Size, allBold)))
                        Next
                    Next
                    sb.AppendLine(ind & Fmt("<p{0} style=""{1}"">{2}</p>", If(allBold, " class=""sub""", ""), st.ToString(), parts.ToString()))
                    Return
                End If
                If n.Indent > 0.7 Then
                    st.Append(Fmt(If(n.Lines.Count = 1, "margin-left:{0}mm;", "text-indent:{0}mm;"), R2(n.Indent)))
                End If
                Dim texts = n.Lines.Select(Function(l) InlineHtml(l, ctx.Size, allBold))
                sb.AppendLine(ind & Fmt("<p{0}{1}>{2}</p>", If(allBold, " class=""sub""", ""),
                                      If(st.Length > 0, " style=""" & st.ToString() & """", ""),
                                      String.Join(" ", texts)))
        End Select
    End Sub

    Private Function MarginAttr(mm As Double) As String
        Return If(mm > 0.05, Fmt(" style=""margin-top:{0}mm""", R2(mm)), "")
    End Function

    ' ------------------------------------------------------------------
    '  Inline: segmen → teks + <b>/<span font>/<span color>
    ' ------------------------------------------------------------------
    Private Function InlineHtml(l As Line, baseSize As Double, boldContext As Boolean) As String
        Return InlineSegs(l.Segs, baseSize, boldContext)
    End Function

    Private Function InlineSegs(segs As List(Of TextRun), baseSize As Double, boldContext As Boolean) As String
        Dim sb As New StringBuilder()
        Dim prevRight As Double = Double.NaN
        For Each s In segs
            If Not Double.IsNaN(prevRight) Then
                Dim gap = s.XMm - prevRight
                If gap > 0.25 * GlobalSettings.SpaceWidthEm * s.SizePt * MmPerPt Then sb.Append(" ")
            End If
            Dim txt = WebUtility.HtmlEncode(s.Text)
            Dim st As New StringBuilder()
            If Math.Abs(s.SizePt - baseSize) > 0.2 Then st.Append(Fmt("font-size:{0}pt;", R2(s.SizePt)))
            If GlobalSettings.UsePdfTextColor OrElse Luminance(s.ColorHex) >= 0.3 OrElse IsColorful(s.ColorHex) Then
                st.Append("color:" & s.ColorHex & ";")
            End If
            If s.IsBold AndAlso Not boldContext Then txt = "<b>" & txt & "</b>"
            If s.IsSuper Then
                ' superscript: ukuran kecil, dinaikkan; line-height:0 supaya tinggi baris tidak berubah
                txt = "<sup style=""" & st.ToString() & "line-height:0"">" & txt & "</sup>"
            ElseIf st.Length > 0 Then
                txt = "<span style=""" & st.ToString() & """>" & txt & "</span>"
            End If
            If s.IsItalic Then txt = "<span class=""it"">" & txt & "</span>"   ' rule: italic PDF → FontNameItalic, style normal
            sb.Append(txt)
            prevRight = s.XMm + s.WidthMm
        Next
        Return sb.ToString()
    End Function

    ''' <summary>Pecah baris di kata ":" (bisa di tengah run) → key / colon / value sebagai run terpisah.</summary>
    Private Sub SplitAtColon(l As Line, ByRef keySegs As List(Of TextRun), ByRef colon As TextRun, ByRef valSegs As List(Of TextRun))
        keySegs = New List(Of TextRun) : valSegs = New List(Of TextRun) : colon = Nothing
        For Each sg0 In l.Segs
            Dim sg = SplitTrailingColon(sg0)
            If colon IsNot Nothing Then valSegs.Add(sg) : Continue For
            Dim ci = sg.Words.FindIndex(Function(w) w.Text = ":")
            If ci < 0 Then keySegs.Add(sg) : Continue For
            If ci > 0 Then keySegs.Add(SubRun(sg, 0, ci))
            colon = SubRun(sg, ci, ci + 1)
            If ci + 1 < sg.Words.Count Then valSegs.Add(SubRun(sg, ci + 1, sg.Words.Count))
        Next
    End Sub

    ' "Pemasar:" (titik dua menempel di kata) → dua kata: "Pemasar" dan ":"
    Private Function SplitTrailingColon(sg As TextRun) As TextRun
        Dim idx = sg.Words.FindIndex(Function(w) w.Text.Length > 1 AndAlso w.Text.EndsWith(":"))
        If idx < 0 Then Return sg
        Dim c As New TextRun With {.XMm = sg.XMm, .WidthMm = sg.WidthMm, .BaselineMm = sg.BaselineMm, .SizePt = sg.SizePt,
                                   .PdfFontName = sg.PdfFontName, .IsBold = sg.IsBold, .IsItalic = sg.IsItalic, .ColorHex = sg.ColorHex}
        Dim colonW = 0.278 * sg.SizePt * MmPerPt
        For i = 0 To sg.Words.Count - 1
            Dim w = sg.Words(i)
            If i = idx Then
                c.Words.Add(New WordItem With {.Text = w.Text.Substring(0, w.Text.Length - 1), .XMm = w.XMm, .WidthMm = Math.Max(0.1, w.WidthMm - colonW)})
                c.Words.Add(New WordItem With {.Text = ":", .XMm = w.XMm + w.WidthMm - colonW, .WidthMm = colonW})
            Else
                c.Words.Add(w)
            End If
        Next
        Return c
    End Function

    Private Function SubRun(sg As TextRun, fromIdx As Integer, toIdx As Integer) As TextRun
        Dim r As New TextRun With {.BaselineMm = sg.BaselineMm, .SizePt = sg.SizePt, .PdfFontName = sg.PdfFontName,
                                   .IsBold = sg.IsBold, .IsItalic = sg.IsItalic, .ColorHex = sg.ColorHex}
        For i = fromIdx To toIdx - 1 : r.Words.Add(sg.Words(i)) : Next
        r.XMm = r.Words(0).XMm
        r.WidthMm = r.Words.Last().XMm + r.Words.Last().WidthMm - r.XMm
        Return r
    End Function

    ''' <summary>Titik tengah shape ada di dalam kotak?</summary>
    Private Function ShapeInside(s As Shape, b As Shape) As Boolean
        Dim cx = s.XMm + s.WidthMm / 2, cy = s.YMm + s.HeightMm / 2
        Return cx >= b.XMm AndAlso cx <= b.XMm + b.WidthMm AndAlso cy >= b.YMm AndAlso cy <= b.YMm + b.HeightMm
    End Function

    ''' <summary>Titik tengah vertikal run (mm).</summary>
    Private Function RunCenterY(x As TextRun) As Double
        Return x.BaselineMm - x.SizePt * MmPerPt * 0.35
    End Function

    ''' <summary>Titik tengah run ada di dalam kotak?</summary>
    Private Function RunInside(x As TextRun, b As Shape) As Boolean
        Dim cx = x.XMm + x.WidthMm / 2
        Dim cy = RunCenterY(x)
        Return cx >= b.XMm - 0.3 AndAlso cx <= b.XMm + b.WidthMm + 0.3 AndAlso cy >= b.YMm - 0.3 AndAlso cy <= b.YMm + b.HeightMm + 0.3
    End Function

    ' ------------------------------------------------------------------
    '  Helpers geometri
    ' ------------------------------------------------------------------
    Private Function BuildLines(runs As List(Of TextRun)) As List(Of Line)
        Dim lines As New List(Of Line)
        For Each run In runs.OrderBy(Function(x) x.BaselineMm).ThenBy(Function(x) x.XMm)
            Dim ln = lines.LastOrDefault(Function(x) Math.Abs(x.Baseline - run.BaselineMm) <= 0.35)
            If ln Is Nothing Then
                ln = New Line With {.Baseline = run.BaselineMm, .Left = run.XMm, .Right = run.XMm + run.WidthMm, .Size = run.SizePt}
                lines.Add(ln)
            End If
            ln.Segs.Add(run)
            ln.Left = Math.Min(ln.Left, run.XMm)
            ln.Right = Math.Max(ln.Right, run.XMm + run.WidthMm)
            ln.Size = Math.Max(ln.Size, run.SizePt)
        Next
        For Each ln In lines : ln.Segs = ln.Segs.OrderBy(Function(s) s.XMm).ToList() : Next
        lines = lines.OrderBy(Function(x) x.Baseline).ToList()
        AttachSuperscripts(lines)
        MergeMarkerLines(lines)
        Return lines
    End Function

    ''' <summary>
    ''' Superscript: segmen kecil (≤ 0.75× ukuran baris tetangga) yang baseline-nya sedikit di atas
    ''' baseline baris di bawahnya (≤ 0.6 em) dan bersebelahan secara horizontal dengan teks baris itu
    ''' → dipindah ke baris tersebut sebagai segmen IsSuper (dirender &lt;sup&gt;).
    ''' </summary>
    Private Sub AttachSuperscripts(lines As List(Of Line))
        For i = lines.Count - 1 To 0 Step -1
            Dim src = lines(i)
            For Each s In src.Segs.ToList()
                Dim target As Line = Nothing
                For Each t In lines
                    If t Is src Then Continue For
                    Dim mains = t.Segs.Where(Function(x) Not x.IsSuper).ToList()
                    If mains.Count = 0 Then Continue For
                    Dim mainSize = mains.Max(Function(x) x.SizePt)
                    If s.SizePt > 0.75 * mainSize Then Continue For
                    Dim dy = t.Baseline - s.BaselineMm
                    If dy <= 0.35 OrElse dy > 0.6 * mainSize * MmPerPt Then Continue For
                    ' bersebelahan: ada segmen t yang tepinya ≤ 1.5 mm dari s, atau s berada di antara segmen t
                    Dim near = mains.Any(Function(x) Math.Abs(x.XMm + x.WidthMm - s.XMm) <= 1.5 OrElse Math.Abs(s.XMm + s.WidthMm - x.XMm) <= 1.5)
                    If Not near AndAlso Not (s.XMm >= t.Left AndAlso s.XMm + s.WidthMm <= t.Right) Then Continue For
                    If target Is Nothing OrElse Math.Abs(t.Baseline - s.BaselineMm) < Math.Abs(target.Baseline - s.BaselineMm) Then target = t
                Next
                If target Is Nothing Then Continue For
                s.IsSuper = True
                src.Segs.Remove(s)
                target.Segs.Add(s)
                target.Segs = target.Segs.OrderBy(Function(x) x.XMm).ToList()
                target.Left = Math.Min(target.Left, s.XMm) : target.Right = Math.Max(target.Right, s.XMm + s.WidthMm)
            Next
            If src.Segs.Count = 0 Then
                lines.RemoveAt(i)
            Else
                src.Left = src.Segs.Min(Function(x) x.XMm) : src.Right = src.Segs.Max(Function(x) x.XMm + x.WidthMm)
                src.Size = src.Segs.Max(Function(x) x.SizePt)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Baris yang hanya berisi marker (mis. "I.") dan baris berikutnya (baseline beda ≤ 1.5 mm, teksnya di kanan marker)
    ''' digabung — baseline bisa bergeser sedikit kalau baris judul memuat superscript.
    ''' </summary>
    Private Sub MergeMarkerLines(lines As List(Of Line))
        For i = lines.Count - 2 To 0 Step -1
            Dim a = lines(i), b = lines(i + 1)
            If a.Segs.Count <> 1 OrElse a.Segs(0).IsSuper OrElse Not IsMarkerText(a.Segs(0).Text) Then Continue For
            If b.Baseline - a.Baseline > 1.5 OrElse b.Left < a.Right + 0.8 Then Continue For
            b.Segs.Insert(0, a.Segs(0))
            b.Left = a.Left : b.Size = Math.Max(a.Size, b.Size)
            lines.RemoveAt(i)
        Next
    End Sub

    ''' <summary>Jarak antar baris (mm): median selisih baseline berurutan yang wajar; fallback 1.32 × ukuran.</summary>
    Private Function LinePitch(ls As List(Of Line), sizePt As Double) As Double
        Dim fallback = 1.32 * sizePt * MmPerPt
        If ls Is Nothing OrElse ls.Count < 2 Then Return fallback
        Dim sorted = ls.Select(Function(l) l.Baseline).Distinct().OrderBy(Function(b) b).ToList()
        Dim diffs As New List(Of Double)
        For i = 1 To sorted.Count - 1
            Dim d = sorted(i) - sorted(i - 1)
            If d > 0.5 AndAlso d < 2.2 * sizePt * MmPerPt Then diffs.Add(d)
        Next
        If diffs.Count = 0 Then Return fallback
        ' modus selisih baseline (dibulatkan 0.1 mm; seri -> yang terbesar)
        Dim groups = diffs.GroupBy(Function(d) Math.Round(d * 10) / 10).ToList()
        Dim modeGrp = groups.OrderByDescending(Function(g) g.Count()).ThenByDescending(Function(g) g.Key).First()
        Dim sizeMm = sizePt * MmPerPt
        ' Kalau modus terlalu lebar untuk satu spasi (> 1.45 em) - mis. deretan baris label:nilai yang berjarak -
        ' pakai selisih terkecil yang wajar untuk satu spasi (1.05-1.45 em); selisih lebar sisanya menjadi margin.
        If modeGrp.Key > 1.45 * sizeMm Then
            Dim natural = groups.Where(Function(g) g.Key >= 1.05 * sizeMm AndAlso g.Key <= 1.45 * sizeMm).
                                 OrderBy(Function(g) g.Key).FirstOrDefault()
            If natural IsNot Nothing Then Return natural.Min()
        End If
        Return modeGrp.Min()
    End Function

    Private Function DominantSize(ls As List(Of Line)) As Double
        If ls.Count = 0 Then Return GlobalSettings.FontSizePt
        Return ls.GroupBy(Function(l) Math.Round(l.Size, 1)).OrderByDescending(Function(g) g.Count()).First().Key
    End Function

    ''' <summary>Posisi atas line-box (mm) untuk baseline & tinggi baris tertentu.</summary>
    Private Function LineTop(baseline As Double, sizePt As Double, lineHmm As Double) As Double
        Dim s = sizePt * MmPerPt
        Return baseline - ((lineHmm - (Ascent + Descent) * s) / 2 + Ascent * s)
    End Function

    Private Function IsMarkerText(t As String) As Boolean
        Return Regex.IsMatch(t, "^(\d{1,3}\.|[A-Za-z]\.|[ivxIVX]{1,4}\.|\(\d{1,3}\)|\d{1,3}\)|[a-zA-Z]\)|[•·\-–])$")
    End Function

    ' bar judul = fill lebar, tinggi 4–12 mm, warnanya gelap (PDF asli: hitam) atau berwarna (PDF hasil cetak HTML: hijau)
    Private Function IsTitleBar(s As Shape) As Boolean
        If s.FillHex Is Nothing Then Return False
        Return s.WidthMm >= GlobalSettings.TitleBarMinWidthMm AndAlso s.HeightMm >= 4 AndAlso s.HeightMm <= 12 AndAlso
               (Luminance(s.FillHex) < 0.25 OrElse IsColorful(s.FillHex))
    End Function

    Private Function Inside(l As Line, b As Shape) As Boolean
        Dim cy = l.Baseline - l.Size * MmPerPt * 0.35
        Return cy >= b.YMm AndAlso cy <= b.YMm + b.HeightMm AndAlso l.Left >= b.XMm - 0.5 AndAlso l.Right <= b.XMm + b.WidthMm + 0.5
    End Function

    Private Function Touch(a As Shape, b As Shape) As Boolean
        Dim e = 0.6
        Return a.XMm - e <= b.XMm + b.WidthMm AndAlso b.XMm - e <= a.XMm + a.WidthMm AndAlso
               a.YMm - e <= b.YMm + b.HeightMm AndAlso b.YMm - e <= a.YMm + a.HeightMm
    End Function

    Private Function BBox(cells As List(Of Shape)) As Shape
        Dim x0 = cells.Min(Function(c) c.XMm), y0 = cells.Min(Function(c) c.YMm)
        Dim x1 = cells.Max(Function(c) c.XMm + c.WidthMm), y1 = cells.Max(Function(c) c.YMm + c.HeightMm)
        Return New Shape With {.XMm = x0, .YMm = y0, .WidthMm = x1 - x0, .HeightMm = y1 - y0}
    End Function

    ''' <summary>Kelompokkan nilai yang berdekatan (toleransi mm) → daftar nilai unik terurut.</summary>
    Private Function Cluster(vals As List(Of Double), tol As Double) As List(Of Double)
        Dim res As New List(Of Double)
        For Each v In vals.OrderBy(Function(x) x)
            If res.Count = 0 OrElse v - res.Last() > tol Then res.Add(v)
        Next
        Return res
    End Function

    Private Function Fmt(pattern As String, ParamArray args() As Object) As String
        Return String.Format(inv, pattern, args)
    End Function

    Private Function R2(v As Double) As String
        Return Math.Round(v, 2).ToString("0.##", inv)
    End Function

    Private Shared Function Luminance(hex As String) As Double
        If hex Is Nothing OrElse hex.Length < 7 Then Return 0
        Dim r = Convert.ToInt32(hex.Substring(1, 2), 16) / 255.0
        Dim g = Convert.ToInt32(hex.Substring(3, 2), 16) / 255.0
        Dim b = Convert.ToInt32(hex.Substring(5, 2), 16) / 255.0
        Return 0.2126 * r + 0.7152 * g + 0.0722 * b
    End Function

    Private Shared Function IsColorful(hex As String) As Boolean
        If hex Is Nothing OrElse hex.Length < 7 Then Return False
        Dim r = Convert.ToInt32(hex.Substring(1, 2), 16)
        Dim g = Convert.ToInt32(hex.Substring(3, 2), 16)
        Dim b = Convert.ToInt32(hex.Substring(5, 2), 16)
        Return Math.Max(r, Math.Max(g, b)) - Math.Min(r, Math.Min(g, b)) > 40
    End Function

End Class
