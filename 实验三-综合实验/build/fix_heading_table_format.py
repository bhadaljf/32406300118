from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from pathlib import Path
PATH = str(Path(__file__).resolve().parent.parent / 'out' / '面向对象分析与设计实验3-软件243_封媛.docx')
doc = Document(PATH)

def set_run_font(run, east, latin, size, bold=False):
    run.font.name = latin
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.get_or_add_rFonts()
    fonts.set(qn('w:eastAsia'), east)
    fonts.set(qn('w:ascii'), latin)
    fonts.set(qn('w:hAnsi'), latin)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def set_paragraph_shading(paragraph, fill=None):
    ppr = paragraph._p.get_or_add_pPr()
    old = ppr.find(qn('w:shd'))
    if old is not None:
        ppr.remove(old)
    if fill:
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:fill'), fill)
        ppr.append(shd)

# 压缩目录样式：目录条目由 Word 域生成，只能改样式定义，避免目录溢出到第二页留出空白页
for st in doc.styles:
    if not st.name.lower().startswith('toc'):
        continue
    pf = st.paragraph_format
    pf.line_spacing = 1.0
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

for p in doc.paragraphs:
    style = p.style.name
    if style == 'Heading 1':
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = Pt(0)
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.keep_with_next = True
        set_paragraph_shading(p, 'D9D9D9')
        for r in p.runs:
            set_run_font(r, '黑体', 'Times New Roman', 14, True)
    elif style == 'Heading 2':
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = Pt(0)
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(9)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.keep_with_next = True
        set_paragraph_shading(p, None)
        for r in p.runs:
            set_run_font(r, '黑体', 'Times New Roman', 12, True)
    elif style == 'Heading 3':
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = Pt(0)
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.keep_with_next = True
        set_paragraph_shading(p, None)
        for r in p.runs:
            set_run_font(r, '黑体', 'Times New Roman', 10.5, True)

def clear_cell_shading(cell):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = tcpr.find(qn('w:shd'))
    if shd is not None:
        tcpr.remove(shd)

def set_cell_margins(cell, top=80, start=100, bottom=80, end=100):
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    mar = tcpr.first_child_found_in('w:tcMar')
    if mar is None:
        mar = OxmlElement('w:tcMar')
        tcpr.append(mar)
    for side, value in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        tag = 'w:' + side
        node = mar.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            mar.append(node)
        node.set(qn('w:w'), str(value))
        node.set(qn('w:type'), 'dxa')

def set_table_borders(table):
    tblpr = table._tbl.tblPr
    borders = tblpr.find(qn('w:tblBorders'))
    if borders is None:
        borders = OxmlElement('w:tblBorders')
        tblpr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        tag = 'w:' + edge
        node = borders.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            borders.append(node)
        node.set(qn('w:val'), 'single')
        node.set(qn('w:sz'), '6')
        node.set(qn('w:space'), '0')
        node.set(qn('w:color'), '000000')

# 首页两个表格属于学校固定模板，保持不变；正文表格从第3个开始。
for table in doc.tables[2:]:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            clear_cell_shading(cell)
            set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.first_line_indent = Pt(0)
                for r in p.runs:
                    set_run_font(r, '宋体', 'Times New Roman', 10.5, ri == 0)
        if ri == 0:
            trpr = row._tr.get_or_add_trPr()
            if trpr.find(qn('w:tblHeader')) is None:
                header = OxmlElement('w:tblHeader')
                header.set(qn('w:val'), 'true')
                trpr.append(header)
            # 表头与紧随其后的行保持同页，避免出现只有表头落在页底的情况
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.keep_with_next = True

doc.save(PATH)
print(PATH)


