# -*- coding: utf-8 -*-
"""
inspect_template.py —— 解析 .docx 模板，输出 docxgen 将自动使用的样式映射
（各级标题 / 正文 / 题注 / 表格的 styleId 及字体字号）以及模板现有大纲。

用途：编写文档前先看清模板提供了哪些固定样式、层级有几级，确认自动探测是否正确；
若探测不符预期，可据此在 DocxBuilder(styles=...) 中显式指定 styleId。

    python inspect_template.py <模板.docx>
"""
import os, sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docxgen import DocxBuilder, _w


def _font_size(rpr):
    font = size = ''
    if rpr is not None:
        rf = rpr.find(_w('rFonts'))
        if rf is not None:
            font = rf.get(_w('eastAsia')) or rf.get(_w('ascii')) or ''
        sz = rpr.find(_w('sz'))
        if sz is not None:
            size = '%spt' % (int(sz.get(_w('val'))) // 2)
    return font, size


def main(path):
    b = DocxBuilder(path, strip_media=False)
    try:
        st = b.style
        sroot = ET.parse(os.path.join(b.word, 'styles.xml')).getroot()
        info = {}
        for s in sroot.findall(_w('style')):
            info[s.get(_w('styleId'))] = _font_size(s.find(_w('rPr')))

        print('== docxgen 将使用的固定样式 ==')
        for lvl, sid in sorted(st['h'].items()):
            f, z = info.get(sid, ('', ''))
            print('  标题 %d  -> styleId=%-6s 字体=%s 字号=%s' % (lvl, sid, f, z))
        for key, label in (('body', '正文'), ('caption', '题注'), ('table', '表格')):
            sid = st.get(key)
            f, z = info.get(sid, ('', ''))
            print('  %s   -> styleId=%-6s 字体=%s 字号=%s' % (label, sid, f, z))

        print('\n== 模板现有大纲（前 80 条标题）==')
        hids = {v: k for k, v in st['h'].items()}
        body = ET.parse(b.doc_path).getroot().find(_w('body'))
        cnt = 0
        for p in body.iter(_w('p')):
            ps = p.find(_w('pPr') + '/' + _w('pStyle'))
            v = ps.get(_w('val')) if ps is not None else None
            if v in hids:
                txt = ''.join(t.text or '' for t in p.iter(_w('t'))).strip()
                print('  ' * hids[v] + '[H%d] %s' % (hids[v], txt[:50]))
                cnt += 1
                if cnt >= 80:
                    break
    finally:
        b.cleanup()


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python inspect_template.py <模板.docx>')
        sys.exit(1)
    main(sys.argv[1])
