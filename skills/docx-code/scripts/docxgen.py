# -*- coding: utf-8 -*-
"""
docxgen —— 基于模板的 OOXML 文档生成引擎（纯 Python 标准库，跨平台）。

核心思想
--------
复制一个"格式标准"的 .docx 模板，完整继承其 styles.xml / numbering.xml / theme /
页眉页脚 / 节设置等全部样式基础设施，只重写 word/document.xml 的正文 <w:body>。
正文段落一律引用模板**已有的固定样式**（pStyle / tblStyle），从而让新文档的
各级标题、正文、图、表格式与模板严格一致——不自造任何样式。

依赖：仅标准库（zipfile / xml.etree / html / re / struct / shutil / tempfile / os）。
无需 python-docx、lxml、pandoc、LibreOffice，可在 Windows / Linux / macOS 直接运行。

典型用法
--------
    from docxgen import DocxBuilder
    with DocxBuilder('模板.docx') as b:
        b.h(1, '功能模块')
        b.p('本章描述各功能模块……')
        b.h(2, '样本数据接入模块')
        b.table(['序号', '格式', '说明'],
                [['1', 'Excel', '按表头解析'], ['2', 'CSV', '文本表格']],
                caption='支持的文件格式')
        # b.image('shot.png', caption='登录界面')   # 可选：嵌入截图
        info = b.save('输出.docx')
        print(info)   # {'headings': {...}, 'paragraphs': N, 'tables': M}

样式探测
--------
构造时自动从 styles.xml 探测：
  - 各级标题：含 <w:outlineLvl> 的段落样式，按 outline 0..N 映射为 level 1..N+1
  - 正文：default 段落样式或名为 Normal 的样式
  - 题注：名称含 caption / 题注 的段落样式（用于图/表题注，SEQ 域自动编号）
  - 表格：表格类样式（如 Normal Table）
探测结果存于 self.style，可在构造时用 styles= 覆盖，或先运行 inspect_template.py 查看。
"""
import html, os, re, struct, shutil, tempfile, zipfile
import xml.etree.ElementTree as ET

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT  = 'http://schemas.openxmlformats.org/package/2006/content-types'
EMU_PER_PX = 9525  # 96 DPI 下 1 像素对应的 EMU

def _w(t): return '{%s}%s' % (W, t)
def esc(s): return html.escape('' if s is None else str(s), quote=False)


class DocxBuilder:
    def __init__(self, template_path, styles=None, strip_media=True, max_img_width_emu=5486400):
        """
        template_path     : 作为样式来源的 .docx 模板路径
        styles            : 覆盖自动探测，形如 {'h': {1:'2',2:'3'}, 'body':'1', 'caption':'12', 'table':'23'}
        strip_media       : True 则清除模板自带图片/嵌入对象（生成纯文字+表格文档，体积小）
        max_img_width_emu : image() 嵌图的最大宽度，默认约 6 英寸
        """
        self.tmp = tempfile.mkdtemp(prefix='docxgen_')
        with zipfile.ZipFile(template_path) as z:
            z.extractall(self.tmp)
        self.word = os.path.join(self.tmp, 'word')
        self.doc_path = os.path.join(self.word, 'document.xml')
        self.rels_path = os.path.join(self.word, '_rels', 'document.xml.rels')
        self.ct_path = os.path.join(self.tmp, '[Content_Types].xml')
        self.max_img_w = max_img_width_emu
        self._parts, self._tbl_no, self._fig_no, self._img_n = [], 0, 0, 0
        self.style = self._detect_styles()
        if styles:
            for k, v in styles.items():
                if v:
                    self.style[k] = v
        if strip_media:
            self._strip_media()

    # ===== 上下文管理 =====
    def __enter__(self): return self
    def __exit__(self, *a): self.cleanup()
    def cleanup(self): shutil.rmtree(self.tmp, ignore_errors=True)

    # ===== 样式自动探测 =====
    def _detect_styles(self):
        root = ET.parse(os.path.join(self.word, 'styles.xml')).getroot()
        heads, body, caption, table = {}, None, None, None
        for s in root.findall(_w('style')):
            sid, typ = s.get(_w('styleId')), s.get(_w('type'))
            nm = s.find(_w('name')); nm = (nm.get(_w('val')) if nm is not None else '') or ''
            low, ppr = nm.lower(), s.find(_w('pPr'))
            if typ == 'paragraph' and ppr is not None:
                o = ppr.find(_w('outlineLvl'))
                if o is not None:
                    heads.setdefault(int(o.get(_w('val'))), sid)
            if typ == 'paragraph' and body is None and (s.get(_w('default')) == '1' or low == 'normal'):
                body = sid
            if typ == 'paragraph' and caption is None and ('caption' in low or '题注' in nm):
                caption = sid
            if typ == 'table' and table is None and ('table' in low or '表' in nm):
                table = sid
        h = {lvl + 1: heads[lvl] for lvl in range(9) if lvl in heads}
        return {'h': h, 'body': body, 'caption': caption, 'table': table}

    # ===== 内容 API（链式） =====
    def h(self, level, text):
        sid = self.style['h'].get(level)
        if not sid:
            raise ValueError('模板无第 %d 级标题样式（可用级别: %s）' % (level, sorted(self.style['h'])))
        self._parts.append('<w:p><w:pPr><w:pStyle w:val="%s"/></w:pPr><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                           % (sid, esc(text)))
        return self

    def p(self, text, style=None):
        sid = style or self.style.get('body')
        ps = '<w:pStyle w:val="%s"/>' % sid if sid else ''
        self._parts.append('<w:p><w:pPr>%s</w:pPr><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ps, esc(text)))
        return self

    def table(self, headers, rows, caption=None, widths=None, total=8828, header_shade='F2F2F2'):
        """表题注（若给）置于表上方，SEQ 自动编号；表头加粗+底纹；全框线；固定列宽。"""
        if caption:
            self._caption('表', caption)
        n = len(headers)
        if not widths:
            widths = [total // n] * n
            widths[-1] = total - sum(widths[:-1])

        def cell(t, i, hd):
            rpr = '<w:rPr><w:b/></w:rPr>' if hd else ''
            shd = '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % header_shade if hd and header_shade else ''
            return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s<w:vAlign w:val="center"/></w:tcPr>'
                    '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:ind w:firstLine="0" w:firstLineChars="0"/>'
                    '<w:jc w:val="center"/></w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:tc>'
                    % (widths[i], shd, rpr, esc(t)))

        def trow(cells, hd):
            return '<w:tr>' + ''.join(cell(c, i, hd) for i, c in enumerate(cells)) + '</w:tr>'

        grid = ''.join('<w:gridCol w:w="%d"/>' % x for x in widths)
        body = trow(headers, True) + ''.join(trow(r, False) for r in rows)
        bl = lambda t: '<w:%s w:val="single" w:sz="4" w:space="0" w:color="auto"/>' % t
        borders = '<w:tblBorders>' + ''.join(bl(x) for x in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) + '</w:tblBorders>'
        ts = '<w:tblStyle w:val="%s"/>' % self.style['table'] if self.style.get('table') else ''
        self._parts.append('<w:tbl><w:tblPr>%s<w:tblW w:w="%d" w:type="dxa"/><w:jc w:val="center"/>%s<w:tblLayout w:type="fixed"/></w:tblPr>'
                           '<w:tblGrid>%s</w:tblGrid>%s</w:tbl>' % (ts, total, borders, grid, body))
        return self

    def image(self, img_path, caption=None, width_emu=None):
        """嵌入本地图片（png/jpg），居中；图题注（若给）置于图下方，SEQ 自动编号。"""
        ext = os.path.splitext(img_path)[1].lower().lstrip('.')
        pw, ph = self._img_size(img_path)
        cx = int(width_emu or min(pw * EMU_PER_PX, self.max_img_w))
        cy = int(cx * ph / pw)
        self._img_n += 1
        name = 'gen_image%d.%s' % (self._img_n, ext)
        os.makedirs(os.path.join(self.word, 'media'), exist_ok=True)
        shutil.copy(img_path, os.path.join(self.word, 'media', name))
        rid = self._add_image_rel('media/' + name)
        self._ensure_ct(ext)
        did = 1000 + self._img_n
        self._parts.append(
            '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
            '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="image%d"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="image%d"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
            % (cx, cy, did, self._img_n, did, self._img_n, rid, cx, cy))
        if caption:
            self._caption('图', caption)
        return self

    def raw(self, xml):
        """高级：直接追加一段合法的 OOXML 段落/表格 XML。"""
        self._parts.append(xml)
        return self

    # ===== 题注（SEQ 域自动编号） =====
    def _caption(self, kind, text):
        if kind == '表':
            self._tbl_no += 1; n = self._tbl_no
        else:
            self._fig_no += 1; n = self._fig_no
        sid = self.style.get('caption')
        ps = '<w:pStyle w:val="%s"/>' % sid if sid else '<w:jc w:val="center"/>'
        self._parts.append(
            '<w:p><w:pPr>%s</w:pPr>'
            '<w:r><w:t xml:space="preserve">%s </w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> SEQ %s \\* ARABIC </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>%d</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (ps, kind, kind, n, esc(text)))

    # ===== 包内基础设施维护 =====
    def _strip_media(self):
        for d in ('media', 'embeddings'):
            p = os.path.join(self.word, d)
            if os.path.isdir(p):
                shutil.rmtree(p)
        ET.register_namespace('', PKG)
        t = ET.parse(self.rels_path); r = t.getroot()
        for rel in list(r):
            if any(k in rel.get('Type', '') for k in ('/image', '/oleObject', '/package')):
                r.remove(rel)
        t.write(self.rels_path, xml_declaration=True, encoding='UTF-8')

    def _add_image_rel(self, target):
        ET.register_namespace('', PKG)
        t = ET.parse(self.rels_path); r = t.getroot()
        ids = {rel.get('Id') for rel in r}
        n = 1
        while 'rId%d' % n in ids:
            n += 1
        rid = 'rId%d' % n
        el = ET.SubElement(r, '{%s}Relationship' % PKG)
        el.set('Id', rid)
        el.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        el.set('Target', target)
        t.write(self.rels_path, xml_declaration=True, encoding='UTF-8')
        return rid

    def _ensure_ct(self, ext):
        ET.register_namespace('', CT)
        t = ET.parse(self.ct_path); r = t.getroot()
        for d in r.findall('{%s}Default' % CT):
            if (d.get('Extension') or '').lower() == ext:
                return
        mime = {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
                'gif': 'image/gif', 'bmp': 'image/bmp', 'emf': 'image/x-emf',
                'tiff': 'image/tiff'}.get(ext, 'application/octet-stream')
        el = ET.Element('{%s}Default' % CT)
        el.set('Extension', ext); el.set('ContentType', mime)
        r.insert(0, el)  # Default 须在 Override 之前
        t.write(self.ct_path, xml_declaration=True, encoding='UTF-8')

    # ===== 图片尺寸（纯标准库读 png/jpg 头） =====
    def _img_size(self, path):
        with open(path, 'rb') as f:
            head = f.read(32)
            if head[:8] == b'\x89PNG\r\n\x1a\n':
                return struct.unpack('>II', head[16:24])
            if head[:2] == b'\xff\xd8':
                f.seek(2)
                while True:
                    b = f.read(1)
                    if not b:
                        break
                    if b != b'\xff':
                        continue
                    m = f.read(1)
                    while m == b'\xff':
                        m = f.read(1)
                    if m in (b'\xc0', b'\xc1', b'\xc2', b'\xc3'):
                        f.read(3)
                        h, w = struct.unpack('>HH', f.read(4))
                        return (w, h)
                    f.read(struct.unpack('>H', f.read(2))[0] - 2)
        return (800, 600)  # 兜底

    # ===== 组装 + 校验 + 打包 =====
    def build_xml(self):
        """返回新的 document.xml 全文（保留模板根命名空间与正文末尾主 sectPr）。"""
        raw = open(self.doc_path, encoding='utf-8').read()
        s = raw.index('<w:body>') + len('<w:body>')
        e = raw.rindex('</w:body>')
        inner = raw[s:e]
        si = inner.rindex('<w:sectPr')
        ei = inner.index('</w:sectPr>', si) + len('</w:sectPr>')
        return raw[:s] + ''.join(self._parts) + inner[si:ei] + raw[e:]

    def validate(self, new=None):
        """校验 XML 合法性与样式引用完整性；返回结构统计 dict。"""
        new = new or self.build_xml()
        ET.fromstring(new)  # XML 合法性
        used = set(re.findall(r'<w:pStyle w:val="([^"]+)"', new)) | set(re.findall(r'<w:tblStyle w:val="([^"]+)"', new))
        have = {x.get(_w('styleId')) for x in ET.parse(os.path.join(self.word, 'styles.xml')).getroot().findall(_w('style'))}
        missing = used - have
        if missing:
            raise ValueError('引用了模板中不存在的样式: %s' % missing)
        body = ET.fromstring(new).find(_w('body'))
        hids = {v: k for k, v in self.style['h'].items()}
        hs, p, t = {}, 0, 0
        for el in body:
            if el.tag == _w('tbl'):
                t += 1
            elif el.tag == _w('p'):
                ps = el.find(_w('pPr') + '/' + _w('pStyle'))
                v = ps.get(_w('val')) if ps is not None else None
                if v in hids:
                    hs[hids[v]] = hs.get(hids[v], 0) + 1
                else:
                    p += 1
        return {'headings': dict(sorted(hs.items())), 'paragraphs': p, 'tables': t}

    def save(self, out_path):
        """校验通过后打包为 .docx；返回结构统计。可重复调用。"""
        new = self.build_xml()
        info = self.validate(new)
        if os.path.exists(out_path):
            os.remove(out_path)
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(self.tmp):
                for f in files:
                    fp = os.path.join(root, f)
                    arc = os.path.relpath(fp, self.tmp).replace(os.sep, '/')  # zip 内统一用 /
                    if arc == 'word/document.xml':
                        z.writestr(arc, new)
                    else:
                        z.write(fp, arc)
        return info
