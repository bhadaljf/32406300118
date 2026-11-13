from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE=Path(__file__).resolve().parent.parent
REF=BASE/'build'/'assets'/'实验一模板参考.docx'
OUT=BASE/'out'/'32406300118.docx'
DIAG=BASE/'diagrams'
doc=Document(REF)

def set_cell(cell,text,bold=False,center=False):
    cell.text=''
    p=cell.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    r=p.add_run(text); r.bold=bold; font(r)
    cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

def font(run,size=10.5,bold=None):
    run.font.name='Times New Roman'; run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体'); run.font.size=Pt(size)
    if bold is not None: run.bold=bold

def shade(cell,fill='D9D9D9'):
    tcPr=cell._tc.get_or_add_tcPr(); shd=tcPr.find(qn('w:shd'))
    if shd is None: shd=OxmlElement('w:shd'); tcPr.append(shd)
    shd.set(qn('w:fill'),fill)

def repeat_header(row):
    trPr=row._tr.get_or_add_trPr(); el=OxmlElement('w:tblHeader'); el.set(qn('w:val'),'true'); trPr.append(el)

def add_p(text='',style=None,bold_lead=None,center=False,indent=True):
    p=doc.add_paragraph(style=style) if style else doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent and not center: p.paragraph_format.first_line_indent=Pt(21)
    p.paragraph_format.line_spacing=1.5
    if bold_lead and text.startswith(bold_lead):
        r=p.add_run(bold_lead); font(r,bold=True); r=p.add_run(text[len(bold_lead):]); font(r)
    else:
        r=p.add_run(text); font(r)
    return p

def disable_heading_numbering(paragraph):
    ppr=paragraph._p.get_or_add_pPr()
    old=ppr.find(qn('w:numPr'))
    if old is not None: ppr.remove(old)
    numPr=OxmlElement('w:numPr')
    ilvl=OxmlElement('w:ilvl'); ilvl.set(qn('w:val'),'0')
    numId=OxmlElement('w:numId'); numId.set(qn('w:val'),'0')
    numPr.append(ilvl); numPr.append(numId); ppr.append(numPr)

def heading(text,level=1):
    style={1:'Heading 1',2:'Heading 2',3:'Heading 3'}.get(level,'Heading 2')
    p=doc.add_paragraph(style=style); p.paragraph_format.keep_with_next=True
    disable_heading_numbering(p)
    r=p.add_run(text); r.font.color.rgb=None
    return p

def table(headers,rows,widths=None):
    t=doc.add_table(rows=1,cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.style=doc.tables[0].style
    for i,h in enumerate(headers): set_cell(t.rows[0].cells[i],h,True,True); shade(t.rows[0].cells[i])
    repeat_header(t.rows[0])
    for row in rows:
        cells=t.add_row().cells
        for i,v in enumerate(row): set_cell(cells[i],str(v),False, i==0 and len(str(v))<12)
    if widths:
        for row in t.rows:
            for i,w in enumerate(widths): row.cells[i].width=Cm(w)
    for row in t.rows:
        trPr=row._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))
    doc.add_paragraph()
    return t

def picture(name,caption,width=15.5):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.keep_with_next=True
    p.add_run().add_picture(str(DIAG/name),width=Cm(width))
    c=doc.add_paragraph(); c.alignment=WD_ALIGN_PARAGRAPH.CENTER; c.paragraph_format.keep_together=True
    r=c.add_run(caption); font(r,10.5)

# 填写首页信息
t1=doc.tables[0]
set_cell(t1.cell(0,1),'计算机科学与网络工程学院',False,True)
set_cell(t1.cell(0,3),'软件工程 软件243',False,True)
set_cell(t1.cell(0,5),'封媛',False,True)
set_cell(t1.cell(0,7),'32406300118',False,True)
set_cell(t1.cell(1,1),'面向对象分析与设计实验',False,True)
# 第 3 行：实验项目名称的值单元格横跨第 1–5 列，指导老师的值在第 7 列
set_cell(t1.cell(2,1),'实验一 需求分析',False,True)
set_cell(t1.cell(2,7),'王健雄',False,True)
t2=doc.tables[1]
for i,v in enumerate(['32406300118','封媛','软件243','独立完成','需求分析、UML建模、文档编写与复核','','']): set_cell(t2.rows[1].cells[i],v,False,True if i<4 else False)
from docx.text.paragraph import Paragraph as _P
_note_ppr=doc.add_paragraph()
t2._tbl.addnext(_note_ppr._p)
_note_p=_P(_note_ppr._p, doc._body)
_note_p.alignment=WD_ALIGN_PARAGRAPH.LEFT
_note_r=_note_p.add_run('说明：本实验为个人独立完成，业务分工表只列本人记录；模板预留的组员空行不作填写。')
font(_note_r,9)

for p in doc.paragraphs:
    if '开课学院及实验室：' in p.text:
        p.text='开课学院及实验室：电子楼416A\t2026年11月13日'
        for r in p.runs: font(r,9)

# 按 V202610 指导书重写实验目的、实验器材、实验内容、要求四栏
from copy import deepcopy

def clone_para(tmpl, text):
    new=deepcopy(tmpl._p)
    runs=new.findall(qn('w:r'))
    for r in runs[1:]: new.remove(r)
    if not runs: raise RuntimeError('模板段落没有文本运行')
    r=runs[0]
    for t in r.findall(qn('w:t'))+r.findall(qn('w:br')): r.remove(t)
    t=OxmlElement('w:t'); t.set(qn('xml:space'),'preserve'); t.text=text; r.append(t)
    return new

paras=list(doc.paragraphs)
i_start=next(i for i,p in enumerate(paras) if p.text.strip().startswith('一、实验目的'))
i_end=next(i for i,p in enumerate(paras) if '丁丁网上书城' in p.text)
front=paras[i_start:i_end]

def first_of(style_name, blocks):
    return next(p for p in blocks if p.style is not None and p.style.name==style_name)

t_head=paras[i_start]
t_body=first_of('Body Text', front)
t_item=[first_of('内容', front[i:]) for i in
        (next(j for j,p in enumerate(front) if p.style.name=='知识点'),
         next(j for j,p in enumerate(front) if p.text.strip().startswith('二、实验器材')),
         next(j for j,p in enumerate(front) if p.text.strip().startswith('三、实验内容')))]
for j,p in enumerate(front):
    if p.text.strip().startswith('四、要求'):
        t_req_head=p
        t_req_item=next(q for q in front[j+1:] if q.style.name=='知识点' and '遵守纪律' in q.text)
        break

blocks=[
 ('h',t_head,'AI 使用声明'),
 ('b',t_body,'本报告对应实验一活动 1（手工建模）。活动 1 的用例划分、类与状态设计由本人确定，未使用 AI 生成模型内容；模型以 PlantUML 文本维护并渲染为报告插图。活动 2 至活动 4 的提示词、AI 原始输出与评估记录，按课程要求随相应阶段另行提交。'),
 ('b',t_body,'使用工具：PlantUML（VS Code 插件，本地渲染）；图形化对照工具按课程要求在活动 2 前选定。'),
 ('h',t_head,'一、实验目的'),
 ('i0',t_item[0],'熟悉基于用例的面向对象需求分析基本方法；'),
 ('i0',t_item[0],'熟练运用文本化建模工具 PlantUML 进行需求描述与分析；'),
 ('i0',t_item[0],'对业务、功能、非功能和接口需求进行静态和动态的分析；'),
 ('i0',t_item[0],'初步建立手工建模、AI 对照、批判评估的工作方式。'),
 ('h',t_head,'二、实验器材'),
 ('i1',t_item[1],'计算机一台；'),
 ('i1',t_item[1],'PlantUML（在线编辑器或 VS Code 插件），以及 StarUML、Enterprise Architect、Visual Paradigm 中任选一款的图形化建模工具；'),
 ('i1',t_item[1],'大语言模型（实验第二环节启用）。'),
 ('h',t_head,'三、实验内容'),
 ('b',t_body,'本实验按手工建模、提示工程、AI 辅助生成、系统性模型评估四个活动依次进行。本报告对应活动 1，活动 2 至活动 4 按课程安排另行提交。'),
 ('i2',t_item[2],'对系统进行简要的可行性分析、业务目标与开发需求描述；'),
 ('i2',t_item[2],'进行整体层面的业务分析，将系统分解为子系统；'),
 ('i2',t_item[2],'选择本人负责的子系统，用 PlantUML 绘制用例图，建立用例字典并进行用例描述；'),
 ('i2',t_item[2],'选择该子系统中的一个用例，建立详尽的静态模型与动态模型，至少包含用例图、类图、状态图、顺序图各一个；'),
 ('i2',t_item[2],'在活动 2 开启前，将手工模型提交入库。'),
 ('h',t_req_head,'四、要求'),
 ('r',t_req_item,'遵守纪律，及时沟通；'),
 ('r',t_req_item,'实验报告在第 12 周前提交。'),
]

anchor=paras[i_end]
for kind,tmpl,text in blocks:
    anchor._p.addprevious(clone_para(tmpl,text))
for p in front:
    p._p.getparent().remove(p._p)
for p in doc.paragraphs:
    if p.style is not None and p.style.name=='知识点':
        p.paragraph_format.keep_with_next=True

# 删除示例系统主体
body=doc._body._element
start=None
for child in list(body):
    txt=''.join(child.itertext())
    if '丁丁网上书城  系统' in txt or '丁丁网上书城 系统' in txt:
        start=child; break
if start is None: raise RuntimeError('未找到示例主体起点')
children=list(body); idx=children.index(start)
for child in children[idx:]:
    if child.tag == qn('w:sectPr'):
        continue
    body.remove(child)

doc.add_page_break()
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(100)
r=p.add_run('校园实验室预约管理系统'); font(r,22,True)
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run('需求分析实验报告'); font(r,20,True)
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(30)
r=p.add_run('Version 1.1'); font(r,14,True)
add_p('作者：封媛　　班级：软件243　　学号：32406300118',center=True,indent=False)
add_p('完成日期：2026年11月13日',center=True,indent=False)

def add_toc_field(paragraph):
    r=paragraph.add_run()
    begin=OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'),'begin')
    instr=OxmlElement('w:instrText'); instr.set(qn('xml:space'),'preserve')
    instr.text='TOC \\o "1-3" \\h \\z \\u'
    sep=OxmlElement('w:fldChar'); sep.set(qn('w:fldCharType'),'separate')
    r._r.append(begin); r._r.append(instr); r._r.append(sep)
    t=OxmlElement('w:t'); t.text='目录将在 Word 中更新'
    hold=paragraph.add_run(); hold._r.append(t)
    end=OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'),'end')
    tail=paragraph.add_run(); tail._r.append(end)

doc.add_page_break(); p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run('目　录'); font(r,18,True)
p=doc.add_paragraph(); p.paragraph_format.first_line_indent=Pt(0); p.paragraph_format.line_spacing=1.5
add_toc_field(p)

h1=heading('1 概述',1); h1.paragraph_format.page_break_before=True
heading('1.1 背景',2)
add_p('学校计算机实验室既承担课程实验，也支持竞赛训练、创新项目和学生自主实践。实验室的开放时间、可容纳人数和设备条件存在差异。当前依赖口头询问或人工登记的预约方式容易出现信息不同步、时段冲突和责任记录不完整等问题。校园实验室预约管理系统将统一维护实验室资源及可预约时段，使申请、审核、签到和使用结果形成可追踪的业务闭环。')
heading('1.2 建设目标',2)
add_p('本系统以减少重复沟通、避免时段冲突和提高实验室利用率为目标。系统应向学生和教师提供清晰的资源查询及预约入口，向实验室管理员提供审核和使用登记依据，并通过统一的状态规则保证预约信息可理解、可核查。')
heading('1.3 系统范围',2)
add_p('系统范围包括用户身份识别、实验室信息查询、可预约时段查询、预约申请、预约审核、取消预约、签到和使用结果登记。课程排课、实验设备借用、耗材采购、门禁硬件控制和财务结算不属于本实验的分析范围。明确边界能够避免把预约系统扩展为完整的实验室综合管理平台。')
heading('1.4 可行性分析',2)
table(['分析维度','结论','依据'],[
['技术可行性','可行','业务以信息查询、规则校验和状态流转为主，可由常见Web应用与关系数据库支持。'],
['操作可行性','可行','参与者角色明确，主要操作与现有预约、审批习惯一致。'],
['管理可行性','可行','审核责任与管理范围可绑定到管理员账号，业务记录能够留痕。'],
['进度可行性','可行','实验一仅完成需求分析和模型，不涉及完整软件实现。']], [3,3,9])
heading('1.5 假设与约束',2)
add_p('系统假设学校已具有统一身份数据，且实验室管理员能够维护开放时间。每个预约申请只对应一个实验室和一个连续时段。申请人数不得超过实验室容量；同一实验室的有效预约不得发生时间重叠。实验一只建立分析模型，不对具体技术框架和数据库产品作强制约束。')
heading('1.6 业务术语',2)
table(['术语','解释'],[
['实验室','可被预约使用的计算机实验场所，具有位置、容量和开放状态等属性。'],
['可预约时段','实验室在特定日期内允许提出预约申请的连续时间范围。'],
['预约申请','申请人针对一个实验室及其时段提出的使用请求。'],
['审核记录','管理员对预约申请给出的审核结果、意见和时间记录。'],
['使用记录','预约通过后形成的签到、离场及使用结果信息。'],
['申请人','提交预约申请的学生或教师，是参与者角色在用例中的实例化称呼。'],
['有效预约','未被取消、驳回或过期，且仍占用时段的预约申请。']], [4,11])

heading('2 业务与功能分析',1)
heading('2.1 参与者分析',2)
table(['参与者','主要目标','业务权限'],[
['学生','获得自主实践或项目活动所需的实验室时段','查询、申请、取消、签到和查看本人预约'],
['教师','为课程补充实验、竞赛辅导或项目活动预约场地','查询、申请、取消、签到和查看本人预约'],
['实验室管理员','保障实验室开放秩序并确认使用情况','维护实验室、审核申请、登记使用结果'],
['系统管理员','维持账号和角色数据的有效性','管理用户、角色与权限']], [3,5,7])
heading('2.2 子系统划分',2)
table(['子系统','职责','核心业务对象','覆盖用例'],[
 ['用户与权限管理','维护用户身份、角色和可访问范围','用户、角色','UC10'],
 ['实验室资源管理','维护实验室及其开放状态与可预约时段','实验室、可预约时段','UC01、UC02、UC09'],
 ['预约业务','受理预约申请，完成审核、签到与使用登记','预约申请、审核记录、使用记录','UC03–UC08']], [3.6,6.2,4.4,3.0])
add_p('三个子系统中，用户与权限管理和实验室资源管理为预约业务提供身份与资源前提。本报告以预约业务子系统为分析对象，用例图、用例字典与四类模型均限定在该子系统范围内；UC01、UC02、UC09、UC10 仅作为外部依赖说明，不展开建模。图 1 中 UC03 对 UC02 的 include 表示业务依赖：提交预约申请必然需要查看可预约时段，而 UC02 属实验室资源管理子系统，其详细建模不在本报告范围。')
heading('2.3 整体业务规则',2)
add_p('系统首先依据用户身份确定其可执行的业务。申请人在查询实验室和可预约时段后提交预约。系统应检查实验室开放状态、人数上限和时段冲突。校验通过后，预约申请进入待审核状态；管理员依据用途、时段和管理要求给出通过或驳回结果。已通过的预约可以签到并最终登记使用结果。取消、驳回和完成均应保留历史记录，不得直接删除。')
picture('01_预约业务子系统用例图.png','图1 预约业务子系统用例图（源文件 01_预约业务子系统用例图.puml）',15.5)

h3=heading('3 用例字典',1); h3.paragraph_format.page_break_before=True
add_p('用例字典覆盖预约业务子系统的六个用例。条目字段包括用例名称、参与者、目标、前置条件、后置条件、基本流与备选流；UC03 的完整描述见 4.1。')
usecases=[
 ['UC03','提交预约申请','学生、教师',
  '目标：提交实验室时段使用申请。\n'
  '前置条件：申请已通过身份认证，实验室与时段信息有效。\n'
  '后置条件：生成状态为待审核的申请，或按失败原因不产生记录。\n'
  '基本流：选择实验室与日期；显示可预约时段；填写用途与人数；校验开放状态、容量与冲突；生成编号并保存。\n'
  '备选流：无可预约时段；人数超过容量；时段冲突；信息不完整；保存失败。详细描述见 4.1。'],
 ['UC04','查询预约记录','学生、教师',
  '目标：查看本人预约及其当前状态。\n'
  '前置条件：申请人已通过身份认证。\n'
  '后置条件：不改变任何业务对象的状态。\n'
  '基本流：进入本人预约列表；按时间或状态筛选；查看实验室、时段与状态。\n'
  '备选流：无记录时提示列表为空；查询他人记录被拒绝。'],
 ['UC05','取消预约','学生、教师',
  '目标：在规则允许范围内终止预约申请。\n'
  '前置条件：申请存在，状态为待审核或已通过，且尚未签到。\n'
  '后置条件：状态变为已取消，保留操作者与操作时间。\n'
  '基本流：选择申请；校验可取消条件；确认取消；置为已取消并留痕。\n'
  '备选流：已签到或已完成不可取消；时段已开始不可取消。'],
 ['UC06','审核预约申请','实验室管理员',
  '目标：对待审核申请给出通过或驳回结论。\n'
  '前置条件：存在待审核申请，管理员具有该实验室的管理范围。\n'
  '后置条件：状态变为已通过或已驳回，生成审核记录。\n'
  '基本流：查看待审核列表；查看申请详情；选择结论并填写意见；更新状态并保存审核记录。\n'
  '备选流：申请已被取消时处理失败；超出管理范围被拒绝。'],
 ['UC07','预约签到','学生、教师',
  '目标：确认申请人按预约到场使用实验室。\n'
  '前置条件：申请已通过且处于签到时间窗口内。\n'
  '后置条件：状态变为已签到，记录签到时间。\n'
  '基本流：申请人发起签到；校验预约与时间窗口；记录签到时间；返回签到结果。\n'
  '备选流：超出时间窗口被拒绝；重复签到给出提示；已取消或已驳回的申请不可签到。'],
 ['UC08','登记使用结果','实验室管理员',
  '目标：登记一次预约的实际使用结果。\n'
  '前置条件：存在已签到的预约。\n'
  '后置条件：状态变为已完成，生成使用记录。\n'
  '基本流：核对在室情况；登记离场时间与使用结果；置为已完成。\n'
  '备选流：未签到的预约不可登记；异常情况写入使用结果备注。']]
table(['用例编号','用例名称','参与者','字典条目'],usecases,[1.9,2.7,2.6,8.4])

heading('4 核心用例分析',1)
heading('4.1 UC03 提交预约申请',2)
table(['项目','内容'],[
['用例编号','UC03'],['用例名称','提交预约申请'],['主要参与者','学生、教师'],['触发条件','申请人需要在指定时间使用某一实验室。'],['前置条件','申请人已通过身份认证；实验室信息及开放时段有效。'],['后置条件','成功时生成状态为“待审核”的预约申请；失败时不产生申请记录。'],['基本流程','1.申请人选择实验室和日期；\n2.系统显示可预约时段；\n3.申请人选择时段并填写用途、人数；\n4.系统校验开放状态、容量和时段冲突；\n5.系统生成预约编号并保存待审核申请；\n6.系统向申请人显示提交结果。'],['扩展流程','2a.没有可预约时段，系统说明原因并结束用例；\n4a.人数超过容量，系统提示修改人数；\n4b.时段已被占用，系统要求重新选择；\n4c.申请信息不完整，系统标明缺失项；\n5a.保存失败，系统不生成预约编号并提示稍后重试。'],['业务规则','同一实验室的有效预约不得时间重叠；申请人数不得超过容量；申请时间必须处于开放范围。']],[3.5,12])
heading('4.2 分析类模型',2)
add_p('分析类模型围绕预约申请建立。用户提出预约，实验室提供资源属性，可预约时段描述时间约束，审核记录保存管理决策，使用记录描述预约履行结果。类之间的关联表达业务关系，不等同于数据库表之间的直接映射。')
add_p('类间关联按业务事实逐条确定。下表按“左端 : 右端”的顺序给出多重性，右端为关联箭头所指的一端。')
add_p('预约申请与可预约时段之间另有直接关联：一次申请占用一个具体时段，一个时段可被多次申请。缺少这条关联时，顺序图中“查询时段状态”的接收对象在类图里找不到落点。')
table(['关联','多重性','业务依据'],[
 ['用户—预约申请','1 : 0..*','一个申请人可提交多次预约；每次申请只属于一个申请人。'],
 ['实验室—预约申请','1 : 0..*','一个实验室可被多次预约；每次申请只针对一个实验室。'],
 ['实验室—可预约时段','1 : 0..*','一个实验室按日期开放多个时段；时段不跨实验室复用。'],
 ['可预约时段—预约申请','1 : 0..*','一个时段可被多次申请；每次申请只占用一个时段。'],
 ['预约申请—审核记录','1 : 0..1','一次申请至多对应一条审核结论；待审核之前不存在审核记录。'],
 ['预约申请—使用记录','1 : 0..1','只有通过的申请才会产生使用记录，且只产生一条。'],
 ['管理员—审核记录','1 : 0..*','一名管理员可处理多条审核记录；每条记录只归属一名管理员。']], [3.6,2.2,9.8])
add_p('审核记录与使用记录采用关联而非组合：两者需要独立留痕，即使预约申请进入终止状态，记录本身仍要可查询，因此不随申请对象级联销毁。类图也未在用户与管理员之间建立泛化，原因是本子系统只关注预约业务所需的身份与权限视图，两者的差异留在用户与权限管理子系统中表达。')
picture('02_预约业务分析类图.png','图2 预约业务子系统分析类图（源文件 02_预约业务分析类图.puml）',15.5)

heading('4.3 预约申请状态模型',2)
add_p('预约申请是本系统中状态变化最完整的业务对象。申请提交后直接进入待审核状态，系统不保存草稿，记录一旦创建即进入业务流程；审核结果将其转为已通过或已驳回。已通过申请可在规定时间签到，使用结束后进入已完成状态。申请人或管理员在规则允许时可以取消申请。状态图另设已过期状态，用于表达申请已通过但未在签到窗口内到场的情形；该状态不再接受签到，也不进入使用登记，从而避免已失效的预约长期停留在已通过。已驳回、已取消、已过期与已完成四个终止状态均保留在系统中用于追踪，不做物理删除；已过期的申请不接受取消操作。到场签到迁移附加守卫条件“在签到窗口内”，窗口之外的签到请求不改变申请状态。')
picture('03_预约申请状态图.png','图3 预约申请状态图（源文件 03_预约申请状态图.puml）',15.5)
heading('4.4 提交预约申请顺序模型',2)
add_p('顺序模型说明UC03基本流程中各分析对象的协作。预约界面负责接收业务信息，预约控制器组织规则校验，可预约时段提供可用状态，预约申请对象保存通过校验的数据。该划分使业务规则与界面表现保持分离。其中预约界面与预约控制器分别是边界类与控制类，按分析类图的惯例不在类图中展开；可预约时段与预约申请两类对象均出现在图 2 中。')
picture('04_提交预约申请顺序图.png','图4 提交预约申请顺序图（源文件 04_提交预约申请顺序图.puml）',12.8)
heading('4.5 模型一致性说明',2)
add_p('四类模型围绕 UC03 展开。下表按用例基本流程逐条列出顺序图中的消息、接收对象与状态迁移，用于核对同一业务事实在不同模型中的表达是否可追溯。')
table(['用例步骤','顺序图消息与序号','接收对象','状态迁移','对应类操作'],[
 ['1 选择实验室和日期','1 选择实验室和时段','预约界面','无，申请尚未创建','—'],
 ['2 系统显示可预约时段','3 查询时段状态；4 返回可用状态','可预约时段','无','可预约时段.检查可用性()'],
 ['3 选择时段并填写用途、人数','2 提交预约信息','预约控制器','无','—'],
 ['4 校验开放状态、容量与冲突','5 校验人数与开放规则','预约控制器','待审核生成前的校验','可预约时段.检查可用性()'],
 ['5 生成编号并保存待审核申请','6 创建预约申请；7 返回预约编号','预约申请','（初始）→待审核','预约申请.提交()'],
 ['6 向申请人显示提交结果','8 返回提交结果；9 显示待审核预约','预约界面','保持待审核','—'],
 ['4a 时段不可用或人数超限','异常分支注释','预约控制器','不迁移，不创建记录','—']], [3.2,4.4,2.6,3.2,2.6])
add_p('核对结论有三点。其一，用例图中的 UC03 与顺序图的消息序列一一对应，未出现用例描述中没有的交互；其二，顺序图的最后一个业务动作落在待审核状态，与状态图中“草稿→待审核”的迁移一致；其三，类图中的预约申请承担状态保存职责，其状态变化由状态图统一说明。四类模型没有引入无法由其他模型解释的核心对象或行为。')

heading('5 非功能性需求',1)
table(['类别','需求','验证方式'],[
 ['正确性','系统应阻止同一实验室的有效预约时间重叠，并保证状态转换符合既定规则。','对同一时段并发提交同一实验室的预约，应仅一条通过；按图 3 的迁移逐条构造用例，未列出的迁移应被拒绝。'],
 ['性能','在正常校园网络环境下，普通查询和提交操作的响应时间宜不超过3秒。','在校园网环境对查询与提交各采样 20 次，统计 95 分位响应时间。'],
 ['安全性','用户只能访问本人预约；管理员只能处理其管理范围内的实验室；关键操作应留存操作者与时间。','以他人账号访问预约记录、以管理员账号处理范围外实验室，均应被拒绝；抽查审核与取消日志的操作者与时间字段。'],
 ['可靠性','预约提交、审核和取消操作应具备事务一致性，不得产生无编号或状态不完整的记录。','在保存环节模拟失败，检查不产生孤立记录；对业务记录逐个核对编号与状态字段是否完整。'],
 ['易用性','系统应明确显示实验室容量、开放状态、可预约时段和失败原因。','按第 7 章验收场景逐条执行，检查提示是否给出可操作的失败原因。'],
 ['可维护性','业务规则、角色权限和状态名称应集中定义，避免在不同模块中出现相互矛盾的解释。','检查规则、角色与状态名称是否只有一处定义，其余位置均为引用。']], [2.4,6.6,6.6])

heading('6 接口与数据需求',1)
heading('6.1 外部接口需求',2)
add_p('系统应从学校统一身份服务获得用户编号、姓名、角色与所属院系四项基础数据，用于确定参与者身份与可访问范围。身份数据在登录时获取一次，会话期间复用，会话超时后重新获取。若身份服务不可用，系统应拒绝预约提交、审核与取消等写操作，仅允许已登录用户对本人预约的只读查询，并在界面说明原因。若后续接入门禁设备，门禁接口只能依据已通过且处于有效时间窗口的预约授权，不应改变预约业务本身的状态规则。本实验暂不规定接口协议与厂商实现。')
heading('6.2 核心数据需求',2)
table(['数据实体','关键数据项','完整性要求'],[
['实验室','编号、名称、位置、容量、开放状态','编号唯一，容量为正整数'],['可预约时段','实验室编号、日期、开始时间、结束时间、状态','开始时间早于结束时间'],['预约申请','预约编号、申请人、实验室、时段、用途、人数、状态','编号唯一，人数不得超过容量'],['审核记录','预约编号、审核人、结果、意见、审核时间','每次审核均可追踪'],['使用记录','预约编号、签到时间、离场时间、使用结果','只能关联已通过的预约']], [3.5,7,5])

heading('7 验收场景',1)
table(['场景','输入或前提','预期结果'],[
['正常提交','实验室开放、时段空闲、人数未超限','生成唯一预约编号，状态为待审核'],['容量超限','申请人数大于实验室容量','拒绝提交并提示修改人数'],['时段冲突','同一实验室已有有效预约覆盖该时段','拒绝提交并要求重新选择'],['审核通过','管理员处理待审核申请并选择通过','申请状态变为已通过并生成审核记录'],['越权访问','用户查询他人预约或管理员跨管理范围审核','系统拒绝操作并记录安全事件'],['完成使用','已通过预约签到并由管理员登记使用结果','状态依次变为已签到、已完成'],
 ['重复签到','同一预约已签到后再次发起签到','拒绝并提示已签到，状态不变'],
 ['已取消申请签到','已取消或已驳回的申请发起签到','拒绝签到，状态不变'],
 ['身份服务不可用','统一身份服务超时或不可用','拒绝写操作，仅保留本人预约的只读查询并提示原因']], [3,7,6])

heading('8 实验中遇到的困难及解决方法',1)
add_p('第一，系统边界容易扩张。初次分析把课程排课、设备借用、门禁控制一并列入范围，参与者与用例数量随之失控。处理办法是先确定“一次预约从申请到使用结束”的业务闭环，再以该闭环判断某项业务是否属于本系统；门禁只作为已通过预约的执行端，不进入分析模型。这一取舍见 1.3 的范围界定与 6.1 对门禁接口的限定。')
add_p('第二，子系统划分粒度反复调整。初版按申请、审核、签到拆成三个子系统，同一业务对象在三个子系统之间来回引用，类图无法收敛。改为按业务对象聚合后，预约申请、审核记录与使用记录归入同一子系统，子系统边界才与状态图的范围对齐，调整结果见 2.2 的子系统表与 4.2 的关联依据表。')
add_p('第三，状态图与活动流程容易混淆。起初把填写表单、点击提交写成状态节点，评审时无法回答“该申请此刻处于什么状态”。改为只保留业务对象可观测的状态，并把触发事件与守卫条件写在迁移上，例如到场签到须处于签到窗口内，窗口之外转入已过期。修正结果见 4.3 图 3，图中只保留业务对象可观测的状态。')
add_p('第四，顺序图容易掺入界面细节。建模时先从用例基本流程逐条提取业务消息，再按边界对象、控制对象和实体对象分配职责，删去提示框、刷新列表一类描述，使顺序图只表达对象之间的协作。对应成果见 4.4 图 4 与 4.5 追溯表，九条消息均可回溯到基本流程。')
add_p('第五，跨图术语不一致。以用例字典为基准逐项核对参与者、对象、状态与业务条件的名称，统一“预约申请”“可预约时段”等术语，并在 4.5 建立追溯表，使每一处不一致都能定位到具体用例步骤；术语的统一口径见 1.6。')

heading('9 心得体会',1)
add_p('本实验把需求分析从功能清单变为可核对的业务模型。决定模型质量的不是图形是否整齐，而是每个建模决策能否给出理由：子系统为什么这样划分、关联为什么取这个多重性、某个状态为什么必须存在。4.2 与 4.3 中的依据表补的正是这部分论据。')
add_p('四类模型各自回答不同问题。用例图确定系统为谁提供什么价值，类图说明承担业务信息与行为的对象，状态图约束单个对象的生命周期，顺序图规定对象之间的协作次序。它们必须共同解释同一个业务事实，因此需要以用例字典为基准反复回核。')
add_p('采用文本化建模之后，模型的修改可以按行比较，一次术语调整的影响范围能在提交记录中直接看到，这比在图形工具中逐个拖动对象更容易发现遗漏。后续进入系统设计阶段，还需在当前分析模型基础上补充体系结构、数据持久化与接口设计，不应把分析类图直接当作程序类结构。')

heading('10 两种建模方式的比较',1)
add_p('课程要求以 PlantUML 为主工具，并另选一款图形化建模工具作对照。两种方式在下列维度上存在差异。')
table(['比较维度','PlantUML 文本化建模','图形化建模工具'],[
 ['模型载体','纯文本 .puml 文件，任何编辑器可读','专用工程文件，需对应工具打开'],
 ['修改与评审','按行 diff，术语可全局替换，提交信息能说明设计决策','拖拽调整直观，但改动无法逐行评审'],
 ['版本管理','多人改同一图可按行合并','文件整体替换，冲突难以合并'],
 ['布局控制','自动布局，复杂图需拆图或调 skinparam','手工摆放，可精确控制版面但调整耗时'],
 ['适用场景','模型反复迭代、需要评审与留痕','一次性出图、版面要求高']], [3.0,6.2,6.4])
add_p('本版四张图均以 PlantUML 完成。用例图与类图改动最频繁，文本化使每次调整都能在提交记录中定位到具体行；状态图改用横向布局后由工具自动排布，省去手工连线；顺序图的对象次序与消息编号直接写在文本里，与用例基本流程逐条对应。代价是复杂图的版面无法精确控制，需要在拆图与调参之间取舍。')

heading('参考文献',1)
add_p('[1] 王健雄, 冯元勇. 面向对象分析与设计实验指导书 V202610[Z]. 广州大学计算机科学与网络工程学院, 2026.',indent=False)
add_p('[2] Object Management Group. OMG Unified Modeling Language Version 2.5.1[EB/OL]. https://www.omg.org/spec/UML/2.5.1, 2017[2026-11-13].',indent=False)
add_p('[3] PlantUML. PlantUML Language Reference Guide[EB/OL]. https://plantuml.com/, 2026[2026-11-13].',indent=False)

happ=heading('附录 A  PlantUML 源文件清单',1); happ.paragraph_format.page_break_before=True
add_p('四个模型源文件随仓库提交，仓库名为学号 32406300118，路径为 diagrams/；报告中的图 1 至图 4 均由下列源文件渲染。')
add_p('渲染命令为 java -jar plantuml.jar -charset UTF-8 -tpng，可用 build/render_diagrams.ps1 一键重建；状态图与用例图在后续提交中有修订，修订说明见提交记录。')
table(['源文件','对应插图','内容摘要','首次提交'],[
 ['01_预约业务子系统用例图.puml','图 1','子系统边界、参与者、六个用例','6fa7c29'],
 ['02_预约业务分析类图.puml','图 2','分析类、关联与多重性','6fa7c29'],
 ['03_预约申请状态图.puml','图 3','状态划分、迁移与守卫条件','6fa7c29'],
 ['04_提交预约申请顺序图.puml','图 4','UC03 消息序列与异常分支','6fa7c29'],
], [7.0,2.0,4.8,2.0])


# 统一新增正文的中英文字体
for p in doc.paragraphs:
    for r in p.runs:
        if r.font.size is None: font(r)

while doc.paragraphs and not doc.paragraphs[-1].text.strip() and not doc.paragraphs[-1].runs:
    last = doc.paragraphs[-1]
    last._p.getparent().remove(last._p)

doc.core_properties.title='校园实验室预约管理系统需求分析实验报告'
doc.core_properties.author='封媛'
doc.core_properties.subject='面向对象分析与设计实验一'
doc.save(OUT)
print(OUT)
