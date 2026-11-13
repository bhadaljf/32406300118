import sys
from docx import Document
from docx.oxml.ns import qn

path = sys.argv[1]
d = Document(path)
print('== 首页表格 ==')
for ti, t in enumerate(d.tables[:2]):
    print(f'-- table {ti}: {len(t.rows)}x{len(t.columns)}')
    for ri, row in enumerate(t.rows):
        vals = []
        for ci, c in enumerate(row.cells):
            vals.append(f'[{ci}]{c.text.strip()}')
        print('  ', ' '.join(vals))
print('== 前若干段 ==')
for p in d.paragraphs[:40]:
    if p.text.strip():
        print(repr(p.style.name), p.text.strip()[:80])
print('== 目录域/标记 ==')
xml = d.element.xml
print('TOC_MARKER' in xml, 'w:fldSimple' in xml, 'TOC \\o' in xml)
