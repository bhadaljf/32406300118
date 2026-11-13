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
OUT=BASE/'out'/'封媛_32406300118_面向对象实验一.docx'
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

def heading(text,level=1):
    style={1:'Heading 1',2:'Heading 2',3:'Heading 3'}.get(level,'Heading 2')
    p=doc.add_paragraph(style=style); p.paragraph_format.keep_with_next=True
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
set_cell(t1.cell(2,1),'需求分析',False,True)
set_cell(t1.cell(2,5),'冯元勇',False,True)
t2=doc.tables[1]
for i,v in enumerate(['32406300118','封媛','软件243','独立完成','需求分析、UML建模、文档编写与复核','','']): set_cell(t2.rows[1].cells[i],v,False,True if i<4 else False)
for p in doc.paragraphs:
    if '开课学院及实验室：' in p.text:
        p.text='开课学院及实验室：计算机科学与工程实验室【实验室待确认】\t2026年11月2日'
        for r in p.runs: font(r,9)

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
r=p.add_run('Version 1.0'); font(r,14,True)
add_p('作者：封媛　　班级：软件243　　学号：32406300118',center=True,indent=False)
add_p('完成日期：2026年11月2日',center=True,indent=False)

doc.add_page_break(); p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run('目　录'); font(r,18,True)
add_p('TOC_MARKER',indent=False)

doc.add_page_break(); heading('1 概述',1)
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
['使用记录','预约通过后形成的签到、离场及使用结果信息。']], [4,11])

heading('2 业务与功能分析',1)
heading('2.1 参与者分析',2)
table(['参与者','主要目标','业务权限'],[
['学生','获得自主实践或项目活动所需的实验室时段','查询、申请、取消、签到和查看本人预约'],
['教师','为课程补充实验、竞赛辅导或项目活动预约场地','查询、申请、取消、签到和查看本人预约'],
['实验室管理员','保障实验室开放秩序并确认使用情况','维护实验室、审核申请、登记使用结果'],
['系统管理员','维持账号和角色数据的有效性','管理用户、角色与权限']], [3,5,7])
heading('2.2 子系统划分',2)
table(['子系统','职责','核心业务对象'],[
['用户与权限管理','维护用户身份、角色和可访问范围','用户、角色'],
['实验室资源管理','维护实验室及其开放信息','实验室、可预约时段'],
['预约申请管理','提交、查询和取消预约','预约申请'],
['预约审核管理','审核申请并保留处理意见','审核记录'],
['签到与使用管理','记录到场和实际使用结果','使用记录']], [4,7,4])
heading('2.3 整体业务规则',2)
add_p('系统首先依据用户身份确定其可执行的业务。申请人在查询实验室和可预约时段后提交预约。系统应检查实验室开放状态、人数上限和时段冲突。校验通过后，预约申请进入待审核状态；管理员依据用途、时段和管理要求给出通过或驳回结果。已通过的预约可以签到并最终登记使用结果。取消、驳回和完成均应保留历史记录，不得直接删除。')
picture('01_系统用例图.png','图1 校园实验室预约管理系统用例图')

heading('3 用例列表',1)
usecases=[
['UC01','查询实验室','学生、教师','按位置、容量和开放状态查询实验室'],['UC02','查看可预约时段','学生、教师','查看指定实验室可申请的时段'],['UC03','提交预约申请','学生、教师','提交用途、人数、实验室和时段信息'],['UC04','查询预约记录','学生、教师','查看本人预约及其当前状态'],['UC05','取消预约','学生、教师','在规则允许范围内取消预约'],['UC06','审核预约申请','实验室管理员','通过或驳回待审核申请'],['UC07','预约签到','学生、教师','对已通过且到达时间窗口的预约签到'],['UC08','登记使用结果','实验室管理员','记录离场时间和使用情况'],['UC09','维护实验室信息','实验室管理员','维护实验室及开放状态'],['UC10','管理用户与权限','系统管理员','维护用户角色和权限']]
table(['编号','用例名称','参与者','简要说明'],usecases,[2,4,4,7])

heading('4 核心用例分析',1)
heading('4.1 UC03 提交预约申请',2)
table(['项目','内容'],[
['用例编号','UC03'],['用例名称','提交预约申请'],['主要参与者','学生、教师'],['触发条件','申请人需要在指定时间使用某一实验室。'],['前置条件','申请人已通过身份认证；实验室信息及开放时段有效。'],['后置条件','成功时生成状态为“待审核”的预约申请；失败时不产生申请记录。'],['基本流程','1.申请人选择实验室和日期；\n2.系统显示可预约时段；\n3.申请人选择时段并填写用途、人数；\n4.系统校验开放状态、容量和时段冲突；\n5.系统生成预约编号并保存待审核申请；\n6.系统向申请人显示提交结果。'],['扩展流程','2a.没有可预约时段，系统说明原因并结束用例；\n4a.人数超过容量，系统提示修改人数；\n4b.时段已被占用，系统要求重新选择；\n4c.申请信息不完整，系统标明缺失项；\n5a.保存失败，系统不生成预约编号并提示稍后重试。'],['业务规则','同一实验室的有效预约不得重叠；申请人数不得超过容量；申请时间必须处于开放范围。']],[3.5,12])
heading('4.2 分析类模型',2)
add_p('分析类模型围绕预约申请建立。用户提出预约，实验室提供资源属性，可预约时段描述时间约束，审核记录保存管理决策，使用记录描述预约履行结果。类之间的关联表达业务关系，不等同于数据库表之间的直接映射。')
picture('02_分析类图.png','图2 预约申请管理子系统分析类图',16)
heading('4.3 预约申请状态模型',2)
add_p('预约申请是本系统中状态变化最完整的业务对象。草稿提交后进入待审核状态；审核结果将其转为已通过或已驳回。已通过申请可在规定时间签到，使用结束后进入已完成状态。申请人或管理员在规则允许时可以取消申请。终止状态保留在系统中用于追踪。')
picture('03_预约申请状态图.png','图3 预约申请状态图',16)
heading('4.4 提交预约申请顺序模型',2)
add_p('顺序模型说明UC03基本流程中各分析对象的协作。预约界面负责接收业务信息，预约控制器组织规则校验，时段信息提供可用状态，预约申请对象保存通过校验的数据。该划分使业务规则与界面表现保持分离。')
picture('04_提交预约申请顺序图.png','图4 提交预约申请顺序图',16)
heading('4.5 模型一致性说明',2)
add_p('用例图中的“提交预约申请”对应UC03。用例基本流程中的时段检查、规则校验和申请创建均在顺序图中出现。类图中的预约申请承担状态保存职责，其状态变化由状态图统一说明。四类模型围绕同一业务语义展开，未在某一模型中引入无法由其他模型解释的核心对象或行为。')

heading('5 非功能性需求',1)
table(['类别','需求'],[
['正确性','系统应阻止同一实验室有效预约的时段重叠，并保证状态转换符合既定规则。'],['性能','在正常校园网络环境下，普通查询和提交操作的响应时间宜不超过3秒。'],['安全性','用户只能访问本人预约；管理员只能处理其管理范围内的实验室；关键操作应留存操作者与时间。'],['可靠性','预约提交、审核和取消操作应具备事务一致性，不得产生无编号或状态不完整的记录。'],['易用性','系统应明确显示实验室容量、开放状态、可预约时段和失败原因。'],['可维护性','业务规则、角色权限和状态名称应集中定义，避免在不同模块中出现相互矛盾的解释。']], [3.5,12])

heading('6 接口与数据需求',1)
heading('6.1 外部接口需求',2)
add_p('系统应从学校统一身份服务获得用户编号、姓名和角色等基础数据。若后续接入门禁设备，门禁接口只能依据已通过且处于有效时间窗口的预约授权，不应改变预约业务本身的状态规则。本实验暂不规定接口协议和厂商实现。')
heading('6.2 核心数据需求',2)
table(['数据实体','关键数据项','完整性要求'],[
['实验室','编号、名称、位置、容量、开放状态','编号唯一，容量为正整数'],['可预约时段','实验室编号、日期、开始时间、结束时间、状态','开始时间早于结束时间'],['预约申请','预约编号、申请人、实验室、时段、用途、人数、状态','编号唯一，人数不得超过容量'],['审核记录','预约编号、审核人、结果、意见、审核时间','每次审核均可追踪'],['使用记录','预约编号、签到时间、离场时间、使用结果','只能关联已通过的预约']], [3.5,7,5])

heading('7 验收场景',1)
table(['场景','输入或前提','预期结果'],[
['正常提交','实验室开放、时段空闲、人数未超限','生成唯一预约编号，状态为待审核'],['容量超限','申请人数大于实验室容量','拒绝提交并提示修改人数'],['时段冲突','同一实验室已有有效预约覆盖该时段','拒绝提交并要求重新选择'],['审核通过','管理员处理待审核申请并选择通过','申请状态变为已通过并生成审核记录'],['越权访问','用户查询他人预约或管理员跨管理范围审核','系统拒绝操作并记录安全事件'],['完成使用','已通过预约签到并由管理员登记使用结果','状态依次变为已签到、已完成']], [3,7,6])

heading('8 实验中遇到的困难及解决方法',1)
add_p('第一，初始需求容易把“实验室预约”扩展为排课、设备借用和门禁控制等综合管理问题。为控制分析范围，本实验以预约申请的生命周期为主线，明确排除与核心目标无直接关系的业务。')
add_p('第二，状态图和活动流程容易混淆。解决时只选择“预约申请”一个对象，检查每个节点能否回答“该申请此时处于什么状态”，从而避免把填写表单、点击按钮等操作写成状态。')
add_p('第三，顺序图容易掺入界面细节。建模时先从用例基本流程提取业务消息，再按边界对象、控制对象和实体对象划分职责，使顺序图表达协作关系而不是页面操作说明。')
add_p('第四，多个模型可能出现术语不一致。完成各图后，以用例表为基准逐项核对参与者、对象名称、状态名称和业务条件，并统一使用“预约申请”“可预约时段”等术语。')

heading('9 心得体会',1)
add_p('本实验使需求分析从功能清单转变为可验证的业务模型。系统边界决定了分析工作是否可控，用例描述决定了后续模型是否一致。通过围绕“提交预约申请”逐步建立用例、类、状态和顺序模型，可以发现仅靠文字描述不易暴露的时段冲突、容量限制和状态终止条件。')
add_p('状态图强调对象生命周期，顺序图强调对象之间的消息次序，类图则说明承担业务信息和行为的对象。三者用途不同，但必须共同解释同一个业务事实。后续若进入系统设计阶段，还需在当前分析模型基础上补充体系结构、数据持久化和接口设计，而不应直接把分析类图当作最终程序类结构。')

heading('参考文献',1)
add_p('[1] 冯元勇. 面向对象分析与设计实验指导书V20241028[Z]. 广州大学, 2024.',indent=False)
add_p('[2] 冯元勇. 代大作业模板及写作指导[Z]. 2019.',indent=False)
add_p('[3] Object Management Group. OMG Unified Modeling Language Version 2.5.1[EB/OL]. https://www.omg.org/spec/UML/2.5.1, 2017.',indent=False)

# 统一新增正文的中英文字体
for p in doc.paragraphs:
    for r in p.runs:
        if r.font.size is None: font(r)

doc.core_properties.title='校园实验室预约管理系统需求分析实验报告'
doc.core_properties.author='封媛'
doc.core_properties.subject='面向对象分析与设计实验一'
doc.save(OUT)
print(OUT)
