# -*- coding: utf-8 -*-
from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

LAB = Path(__file__).resolve().parent.parent          # 实验二-系统设计
BASE = LAB.parent                                      # 仓库根
REF = LAB / 'build' / 'assets' / '实验二模板参考.docx'
OUT = LAB / 'out' / '面向对象分析与设计实验2-软件243_封媛.docx'
DIAG = LAB / 'diagrams'
doc = Document(REF)


def font(run, size=10.5, bold=None):
    run.font.name = 'Times New Roman'
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold


def set_cell(cell, text, bold=False, center=False):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(text)
    r.bold = bold
    font(r)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def shade(cell, fill='D9D9D9'):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:fill'), fill)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement('w:tblHeader')
    el.set(qn('w:val'), 'true')
    trPr.append(el)


def disable_heading_numbering(paragraph):
    ppr = paragraph._p.get_or_add_pPr()
    old = ppr.find(qn('w:numPr'))
    if old is not None:
        ppr.remove(old)
    numPr = OxmlElement('w:numPr')
    ilvl = OxmlElement('w:ilvl'); ilvl.set(qn('w:val'), '0')
    numId = OxmlElement('w:numId'); numId.set(qn('w:val'), '0')
    numPr.append(ilvl); numPr.append(numId)
    ppr.append(numPr)


def add_p(text='', style=None, center=False, indent=True):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent and not center:
        p.paragraph_format.first_line_indent = Pt(21)
    p.paragraph_format.line_spacing = 1.5
    r = p.add_run(text)
    font(r)
    return p


def heading(text, level=1):
    style = {1: 'Heading 1', 2: 'Heading 2', 3: 'Heading 3'}.get(level, 'Heading 2')
    p = doc.add_paragraph(style=style)
    p.paragraph_format.keep_with_next = True
    disable_heading_numbering(p)
    r = p.add_run(text)
    r.font.color.rgb = None
    return p


def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = doc.tables[0].style
    for i, h in enumerate(headers):
        set_cell(t.rows[0].cells[i], h, True, True)
        shade(t.rows[0].cells[i])
    repeat_header(t.rows[0])
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            set_cell(cells[i], str(v), False, i == 0 and len(str(v)) < 12)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    for row in t.rows:
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))
    for cell in t.rows[0].cells:
        for p in cell.paragraphs:
            p.paragraph_format.keep_with_next = True
    doc.add_paragraph()
    return t


def picture(name, caption, width=15.0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(DIAG / name), width=Cm(width))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.keep_together = True
    r = c.add_run(caption)
    font(r, 10.5)


def add_toc_field(paragraph):
    r = paragraph.add_run()
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = 'TOC \\o "1-3" \\h \\z \\u'
    sep = OxmlElement('w:fldChar'); sep.set(qn('w:fldCharType'), 'separate')
    r._r.append(begin); r._r.append(instr); r._r.append(sep)
    t = OxmlElement('w:t'); t.text = '目录将在 Word 中更新'
    hold = paragraph.add_run(); hold._r.append(t)
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    tail = paragraph.add_run(); tail._r.append(end)


def clone_para(tmpl, text):
    new = deepcopy(tmpl._p)
    runs = new.findall(qn('w:r'))
    for r in runs[1:]:
        new.remove(r)
    if not runs:
        raise RuntimeError('模板段落没有文本运行')
    r = runs[0]
    for t in r.findall(qn('w:t')) + r.findall(qn('w:br')):
        r.remove(t)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    r.append(t)
    return new


# ---------------- 首页 ----------------
t1 = doc.tables[0]
set_cell(t1.cell(0, 1), '计算机科学与网络工程学院', False, True)
set_cell(t1.cell(0, 3), '软件工程 软件243', False, True)
set_cell(t1.cell(0, 5), '封媛', False, True)
set_cell(t1.cell(0, 7), '32406300118', False, True)
set_cell(t1.cell(1, 1), '面向对象分析与设计实验', False, True)
set_cell(t1.cell(2, 1), '实验二 系统设计', False, True)
set_cell(t1.cell(2, 7), '王健雄', False, True)

t2 = doc.tables[1]
for i, v in enumerate(['32406300118', '封媛', '软件243', '独立完成', '体系结构设计、详细设计、模型校验与文档编写', '', '']):
    set_cell(t2.rows[1].cells[i], v, False, True if i < 4 else False)

for p in doc.paragraphs:
    if '开课学院及实验室：' in p.text:
        p.text = '开课学院及实验室：电子楼416A\t2026年11月16日'
        for r in p.runs:
            font(r, 9)

# ---------------- 四栏 ----------------
paras = list(doc.paragraphs)
i_start = next(i for i, p in enumerate(paras) if p.text.strip().startswith('一、实验目的'))
i_end = next(i for i, p in enumerate(paras) if '丁丁网上书城' in p.text)
front = paras[i_start:i_end]


def first_of(style_name, blocks):
    return next(p for p in blocks if p.style is not None and p.style.name == style_name)


t_head = paras[i_start]
t_body = first_of('Body Text', front)
t_item = [first_of('内容', front[i:]) for i in
          (next(j for j, p in enumerate(front) if p.style.name == '知识点'),
           next(j for j, p in enumerate(front) if p.text.strip().startswith('二、实验器材')),
           next(j for j, p in enumerate(front) if p.text.strip().startswith('三、实验内容')))]
for j, p in enumerate(front):
    if p.text.strip().startswith('四、要求'):
        t_req_head = p
        t_req_item = next(q for q in front[j + 1:] if q.style.name == '知识点' and '遵守纪律' in q.text)
        break

blocks = [
 ('h', t_head, 'AI 使用声明'),
 ('b', t_body, '活动 1（手工设计）的体系结构方案、类与接口设计由本人确定，未使用 AI 生成设计内容，模型以 PlantUML 文本维护并渲染为报告插图。活动 2 至活动 4 的提示词、AI 备选方案与论证结论在完成后于第 11 章和附录 B 补充。'),
 ('b', t_body, '使用工具：PlantUML（VS Code 插件，本地渲染）与 Git（本地仓库）。'),
 ('h', t_head, '一、实验目的'),
 ('i0', t_item[0], '熟悉面向对象系统设计的基本方法与过程；'),
 ('i0', t_item[0], '了解并运用 4+1 视图模型描述系统设计；'),
 ('i0', t_item[0], '掌握 diagram-as-code 的工程化工作方式：模型入版本库、以 diff 评审模型变更；'),
 ('i0', t_item[0], '理解 AI 工具在设计方案生成中的能力与局限。'),
 ('h', t_head, '二、实验器材'),
 ('i1', t_item[1], '计算机一台；'),
 ('i1', t_item[1], 'PlantUML（在线编辑器或 VS Code 插件），以及 StarUML、Enterprise Architect、Visual Paradigm 中任选一款的图形化建模工具；'),
 ('i1', t_item[1], 'Git（本地仓库即可）；'),
 ('i1', t_item[1], '大语言模型（实验第二环节启用）。'),
 ('h', t_head, '三、实验内容'),
 ('b', t_body, '本实验按手工设计、提示工程、AI 辅助生成、系统性模型评估与方案论证四个活动依次进行。本报告对应活动 1，并给出活动 2 的提示词设计；活动 3、4 的内容随对照方案一并补充。'),
 ('i2', t_item[2], '基于实验一的系统需求，对整个系统进行体系结构设计；'),
 ('i2', t_item[2], '针对实验一中选择的 UC03 提交预约申请进行详尽设计；'),
 ('i2', t_item[2], '按 4+1 视图给出逻辑模型、开发模型、数据模型、部署模型与场景模型，至少各一张图；'),
 ('i2', t_item[2], '全部模型以 PlantUML 文本提交至仓库，提交信息说明设计决策；'),
 ('i2', t_item[2], '在活动 2 开启前完成手工设计的入库。'),
 ('h', t_req_head, '四、要求'),
 ('r', t_req_item, '遵守纪律，及时沟通；'),
 ('r', t_req_item, '实验报告按课程安排提交。'),
]

anchor = paras[i_end]
for kind, tmpl, text in blocks:
    anchor._p.addprevious(clone_para(tmpl, text))
for p in front:
    p._p.getparent().remove(p._p)
for p in doc.paragraphs:
    if p.style is not None and p.style.name == '知识点':
        p.paragraph_format.keep_with_next = True

# ---------------- 删除模板示例主体 ----------------
body = doc._body._element
start = None
for child in list(body):
    txt = ''.join(child.itertext())
    if '丁丁网上书城  系统' in txt or '丁丁网上书城 系统' in txt:
        start = child
        break
if start is None:
    raise RuntimeError('未找到示例主体起点')
children = list(body)
idx = children.index(start)
for child in children[idx:]:
    if child.tag == qn('w:sectPr'):
        continue
    body.remove(child)

# ---------------- 封面 ----------------
doc.add_page_break()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(100)
r = p.add_run('校园实验室预约管理系统'); font(r, 22, True)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('系统设计实验报告'); font(r, 20, True)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(30)
r = p.add_run('Version 1.0'); font(r, 14, True)
add_p('作者：封媛　　班级：软件243　　学号：32406300118', center=True, indent=False)
add_p('完成日期：2026年11月16日', center=True, indent=False)

# ---------------- 目录 ----------------
h = heading('目录占位', 1)   # 占位后删除
h._p.getparent().remove(h._p)
doc.add_page_break()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('目　录'); font(r, 18, True)
p = doc.add_paragraph(); p.paragraph_format.first_line_indent = Pt(0); p.paragraph_format.line_spacing = 1.5
add_toc_field(p)

h1 = heading('1 设计概述', 1); h1.paragraph_format.page_break_before = True

heading('1.1 与实验一的关系', 2)
add_p('实验一确立了预约业务子系统的六个用例与一套分析模型，本实验在其基础上完成体系结构设计与详细设计。分析模型中的类保持业务语义不变，设计阶段新增分层、接口、数据表与部署节点，不改变用例的边界。')
add_p('两阶段模型的对应关系见第 9 章。凡是分析模型没有的对象，设计阶段也不凭空引入；凡是被拆分的分析类，都在映射表中给出拆分理由。')

heading('1.2 设计目标与约束', 2)
table(['目标或约束', '内容'], [
 ['功能范围', 'UC03 至 UC08：提交、查询、取消、审核、签到、登记使用结果'],
 ['非功能约束', '普通查询与提交响应时间不超过 3 秒；关键操作可追溯；单实验室时段不得冲突'],
 ['运行环境', '校园网内的 Web 应用；浏览器访问；学校统一身份服务提供身份数据'],
 ['数据保留', '审核记录与使用记录独立留痕，不随申请删除'],
 ['实现约束', '课程实验不要求完整实现，设计与原型为准']], [3.4, 12.0])

heading('1.3 设计原则', 2)
add_p('设计遵循三条原则。其一，依赖方向朝内：表现层依赖应用层，应用层依赖领域层与基础设施接口，领域层不依赖任何外层。其二，接口隔离变化：身份校验、持久化、消息通知三类易变能力以接口形式声明，实现可替换。其三，决策留痕：每个非显然的选择都记录备选方案与取舍理由，见第 7 章。')

heading('1.4 4+1 视图与本文档的对应', 2)
add_p('4+1 视图用五个视角描述同一套设计。下表给出本报告的组织方式，避免“画了五张图却说不清哪张回答什么问题”。')
table(['视图', '回答的问题', '本文位置'], [
 ['逻辑视图', '系统由哪些对象构成、如何协作', '第 3 章，图 3'],
 ['开发视图', '代码如何分层组织、层间依赖方向', '2.2 节，图 1'],
 ['数据视图', '业务对象如何落地为表与约束', '第 4 章，图 4'],
 ['部署视图', '组件如何分布到节点', '2.3 节，图 2'],
 ['场景视图', '一次用例在设计中如何执行', '5.1 节，图 5'],
 ['（补充）行为视图', '对象状态如何随时间变化', '5.2 节，图 6'],
 ['（补充）界面视图', '用户如何与系统交互', '第 6 章，图 7']], [3.0, 7.4, 5.6])

# ---------------- 2 体系结构 ----------------
heading('2 体系结构设计', 1)
heading('2.1 风格选择', 2)
add_p('备选三种风格：单体分层、微服务、事件驱动。本系统的规模与部署条件决定了分层单体最合适：业务对象只有五个，跨服务调用带来的分布式事务成本远高于收益；课程环境只提供单台应用服务器，微服务的独立部署无从体现。事件驱动的价值在于审核结果的通知，本设计只把消息通知收敛为一个适配器，保留后续异步化的接口，不整体改为事件架构。')
table(['风格', '适配度', '结论'], [
 ['分层单体', '业务规模小、事务边界清晰、部署简单', '采用'],
 ['微服务', '需要独立伸缩与团队边界，本系统两者都不具备', '不采用'],
 ['事件驱动', '仅通知环节受益，整体改造成本高', '局部预留']], [3.2, 8.6, 3.6])

heading('2.2 分层职责', 2)
table(['层次', '职责', '包含组件'], [
 ['表现层', '接收用户输入、展示结果、做格式校验', '预约界面、审核界面、签到与使用界面'],
 ['应用层', '编排用例流程、控制事务边界、调用领域对象', '预约服务、审核服务、签到与使用服务、身份适配器'],
 ['领域层', '承载业务规则与状态迁移，不依赖框架', '预约申请、可预约时段、审核记录、使用记录'],
 ['基础设施层', '持久化、外部系统接入、消息发送', '预约仓储、记录仓储、消息通知']], [2.6, 8.2, 4.6])
picture('01_分层体系结构.png', '图1 分层体系结构（源文件 01_分层体系结构.puml）', 13.5)

heading('2.3 部署视图', 2)
add_p('系统部署为三个节点：浏览器端单页应用、应用服务器、数据库服务器；身份校验通过学校统一身份服务完成。应用服务器同时承载网关与服务，网关负责鉴权与限流，避免把鉴权逻辑散落到各服务中。')
picture('03_部署图.png', '图2 部署视图（源文件 03_部署图.puml）', 15.0)

# ---------------- 3 逻辑设计 ----------------
heading('3 逻辑设计', 1)
heading('3.1 设计类图', 2)
add_p('设计类图在分析类的基础上补充了控制类、服务类与两类接口。分析阶段的“管理员”类不再作为实体存在，其审核行为上移为审核服务，实体只保留数据与状态迁移——这是设计阶段与分析阶段的主要差别：职责从“谁做”转向“放在哪一层”。')
picture('02_设计类图.png', '图3 设计类图（源文件 02_设计类图.puml）', 15.0)

heading('3.2 关键接口与依赖倒置', 2)
table(['接口', '方法', '引入理由'], [
 ['预约仓储接口', '保存、按编号查询、查冲突', '隔开持久化实现，测试时可替换为内存实现'],
 ['身份接口', '校验令牌', '统一身份服务属外部系统，不能直接依赖其 SDK'],
 ['记录仓储接口', '保存审核记录、保存使用记录', '审核与使用登记共享同一持久化契约']], [3.6, 6.4, 6.0])
add_p('三个接口都定义在应用层，实现放在基础设施层，依赖方向由外向内，符合依赖倒置。这样即使后续把关系数据库换成其他存储，应用层代码不需要改动。')

heading('3.3 分析类到设计类的映射', 2)
table(['分析类', '设计类', '变化'], [
 ['预约申请', '预约申请实体 + 预约状态枚举', '状态由字符串改为枚举，避免非法值'],
 ['可预约时段', '可预约时段实体', '新增冲突判断方法'],
 ['审核记录 / 使用记录', '两类实体 + 记录仓储接口', '抽取统一持久化契约'],
 ['管理员', '审核服务', '行为上移为服务，实体不再承担审核'],
 ['（新增）', '预约服务、预约控制器', '编排用例流程与控制事务边界']], [3.6, 6.2, 6.2])

# ---------------- 4 数据设计 ----------------
heading('4 数据设计', 1)
heading('4.1 数据模型', 2)
add_p('数据模型由五个表构成，主键采用业务编号而非自增序列，便于跨表追溯与日志核对。')
picture('07_数据模型.png', '图4 数据模型（源文件 07_数据模型.puml）', 15.5)

heading('4.2 完整性约束', 2)
table(['约束', '实现方式', '目的'], [
 ['同一实验室时段不重叠', '对（实验室编号，时段编号）建唯一约束 + 提交前查冲突', '既防并发重复，也给出可读的失败原因'],
 ['人数不超过容量', '应用层校验 + 数据库检查约束', '双重保障'],
 ['一次申请至多一条审核记录', '审核记录表对预约编号建唯一索引', '保证审核唯一性'],
 ['使用记录只关联已通过预约', '外键 + 应用层状态校验', '防止对未通过申请登记使用']], [5.0, 6.4, 4.6])

heading('4.3 事务边界', 2)
add_p('三个写操作各自是一个事务：提交预约（校验 + 写入申请）、审核（更新状态 + 写审核记录）、登记使用（更新状态 + 写使用记录）。事务边界落在应用服务的方法上，仓储只负责单表读写，不自行开启事务。归档任务按批提交，单批失败不回滚已归档批次。')

# ---------------- 5 动态设计 ----------------
heading('5 动态设计', 1)
heading('5.1 提交预约的设计级顺序', 2)
add_p('与分析阶段相比，设计级顺序图引入控制器、仓储接口与事务，把“校验”拆成身份校验与业务校验两段，并明确失败时回滚。')
picture('04_提交预约顺序图.png', '图5 提交预约顺序图（设计级，源文件 04_提交预约顺序图.puml）', 12.5)

heading('5.2 状态设计与持久化', 2)
add_p('预约申请的状态集合沿用实验一的设计，未在设计中新增业务状态；已过期仍为终态，只能由归档任务迁移出。状态以枚举持久化，读写时做合法性判断，非法迁移在服务层直接拒绝。')
picture('05_预约申请状态图.png', '图6 预约申请状态图（设计级，源文件 05_预约申请状态图.puml）', 14.0)

heading('5.3 并发与冲突处理', 2)
add_p('两个申请人同时提交同一时段是唯一的实质并发点。设计采用“唯一约束兜底 + 提交前查冲突”的组合：查冲突负责给出友好提示，唯一约束负责在并发穿透时保证不写入重复记录。审核与登记使用按预约编号加行级锁，同一申请的状态变更串行执行。')

# ---------------- 6 界面 ----------------
heading('6 界面设计', 1)
heading('6.1 界面原型', 2)
add_p('界面按用例划分，一屏完成一个用例。提交页把实验室、日期、时段、用途、人数放在同一表单，减少跳转；失败原因就近显示，不弹出独立对话框。')
picture('06_界面原型.png', '图7 提交预约申请界面原型（源文件 06_界面原型.puml）', 8.0)

heading('6.2 交互规则', 2)
table(['场景', '界面行为'], [
 ['时段已被占用', '提交后定位到失败原因，并刷新可选时段'],
 ['人数超过容量', '在人数输入框旁提示上限，不提交请求'],
 ['未登录或令牌过期', '跳转统一身份登录，返回后保留已填内容'],
 ['审核完成', '审核页刷新列表，已处理项移出待办']], [4.2, 11.2])

# ---------------- 7 决策 ----------------
heading('7 设计决策与权衡', 1)
table(['决策', '备选', '取舍理由'], [
 ['分层单体而非微服务', '微服务、事件驱动', '业务对象少、事务边界清晰，分布式成本高于收益'],
 ['实体不承担审核行为', '把审核方法留在管理员实体', '审核涉及记录写入与状态迁移，属跨对象编排'],
 ['仓储以接口声明', '服务直接依赖数据库访问代码', '便于替换实现与单元测试'],
 ['唯一约束 + 查冲突双保险', '仅应用层校验', '并发下应用层校验会穿透'],
 ['状态用枚举持久化', '字符串状态', '避免非法值，迁移判断集中'],
 ['消息通知做成适配器', '直接调用邮件或短信 SDK', '通知渠道可能变化，且非核心链路']], [4.6, 4.8, 6.0])

# ---------------- 8 非功能 ----------------
heading('8 非功能需求在设计中的落实', 1)
table(['需求', '设计手段', '验证方式'], [
 ['响应不超过 3 秒', '列表查询分页；提交路径只做一次冲突查询', '对查询与提交各采样 20 次，取 95 分位'],
 ['关键操作可追溯', '审核与使用记录独立留痕，状态迁移写操作者与时间', '抽查日志与记录表'],
 ['时段不冲突', '唯一约束 + 提交前查冲突', '并发提交同一时段，只允许一条成功'],
 ['身份服务不可用可降级', '身份适配器超时后拒绝写操作，保留只读查询', '断开身份服务后验证读写行为'],
 ['模型可评审', '全部模型为 PlantUML 文本，入仓库并按 diff 评审', '查看提交历史与改动行']], [4.0, 7.0, 5.0])

# ---------------- 9 一致性 ----------------
heading('9 与实验一模型的一致性', 1)
add_p('设计阶段新增了控制器、服务与接口，但没有新增业务概念，也没有删除分析阶段的任何实体。用例与组件的对应关系如下，跨阶段可逐条追溯。')
table(['实验一用例', '设计阶段承载组件', '对应设计图'], [
 ['UC03 提交预约申请', '预约控制器、预约服务、预约仓储接口', '图 3、图 5'],
 ['UC04 查询预约记录', '预约服务（查询方法）', '图 3'],
 ['UC05 取消预约', '预约服务（取消方法）、预约申请', '图 3、图 6'],
 ['UC06 审核预约申请', '审核服务、记录仓储接口', '图 3、图 6'],
 ['UC07 预约签到', '签到与使用服务、预约申请', '图 3、图 6'],
 ['UC08 登记使用结果', '签到与使用服务、使用记录', '图 3、图 6']], [5.2, 7.4, 3.4])

# ---------------- 10 困难 ----------------
heading('10 困难与心得', 1)
add_p('第一，体系结构风格容易一上来就选微服务。真正比较后才发现，本系统只有一个实质并发点、三个写事务，拆成服务只会把简单问题变成分布式一致性问题。选择风格的依据应当是事务边界与团队边界，不是技术热度。')
add_p('第二，分析类不能直接当设计类用。分析阶段把审核行为挂在管理员实体上很自然，设计阶段就必须上移到服务层，否则实体既管数据又管跨表写入，测试时无法单独构造。')
add_p('第三，并发问题不能只靠应用层校验。查冲突与唯一约束是两件事：前者给用户可读的原因，后者保证数据不错。少了哪一半都会留下缺口。')
add_p('第四，4+1 视图不是画五张图交差。逻辑、开发、数据、部署、场景各自回答不同问题，画完后要能用一张图解释另一张图；解释不通，说明某张图画错了。')

# ---------------- 11 提示词 ----------------
heading('11 活动 2 提示词设计', 1)
add_p('活动 2 要求把设计约束编码进提示词，并要求模型给出备选方案而非单一方案。提示词按五要素组织如下。')
table(['要素', '表述'], [
 ['设计任务', '为校园实验室预约管理系统的预约业务子系统给出两套备选体系结构方案，并比较取舍'],
 ['输入基线', '已有六个用例 UC03–UC08 与一套分析类模型，业务对象为预约申请、可预约时段、审核记录、使用记录'],
 ['设计约束', '单台应用服务器；关系数据库；学校统一身份服务；普通操作 3 秒内响应；审核与使用记录独立留痕'],
 ['期望产出', '两套候选方案的 PlantUML 组件图、各自的分层职责表、关键类图，以及对可维护性、可扩展性、需求匹配度的自评'],
 ['禁止项', '不要输出代码实现；不要引入基线中没有的业务概念']], [3.0, 13.0])
add_p('与活动 1 的提示词相比，这一次显式给出了约束与“两套方案”的要求。原因在于设计问题不存在唯一正确答案，只给单一方案会让模型倾向于附和人给出的方向，无法暴露真实取舍。')
add_p('第二轮以“架构评审专家”的角色追加要求，针对动态协作与并发处理深化：职责分离（实体只保留数据与状态迁移，校验编排放到服务）、并发保障（不得只依赖先查后写，须给出唯一约束或锁的兜底）、外部系统隔离（必须引入身份接口），并要求用 alt 片段画出异常分支。这一轮把评审关注点直接写进提示词，是为了看模型在被点名后能否自我修正。')

heading('12 AI 备选方案与论证', 1)
add_p('本章对应活动 3 与活动 4。两轮提示词下的模型输出原样收录，不作修正；评估按“体系结构、职责分配、接口隔离、并发处理、关联语义、跨图一致”六类展开，最后给出三维度论证与选型结论。')

heading('12.1 生成结果', 2)
add_p('第一轮返回两套体系结构方案与一版核心类图；第二轮被要求补充并发分支后，返回设计级顺序图与修正版类图。五张图均为原始输出，代码收录于附录 B。')
picture('ai_A_分层单体.png', '图8 AI 方案A：分层单体组件图（原始输出）', 15.0)
picture('ai_B_事件驱动.png', '图9 AI 方案B：事件驱动组件图（原始输出）', 15.0)
picture('ai_类图_第一轮.png', '图10 AI 核心设计类图（第一轮原始输出）', 15.0)
picture('ai_顺序图_并发.png', '图11 AI 设计级顺序图（含并发分支，原始输出）', 15.0)
picture('ai_类图_第二轮.png', '图12 AI 核心设计类图（修正版，原始输出）', 15.0)

heading('12.2 与手工方案的逐项对照', 2)
table(['检查项', '手工设计', 'AI 输出', '判定'], [
 ['体系结构风格', '分层单体，四层依赖单向朝内', '方案A 同为分层单体，方案B 为事件驱动', 'A 与手工一致；B 属备选'],
 ['依赖方向', '领域层不依赖基础设施实现', '方案A 中“预约申请”指向“预约仓储实现”，依赖倒置画反', '不通过'],
 ['实体职责', '实体只管数据与状态迁移', '第一轮把容量校验、冲突判断放在可预约时段实体；第二轮改到服务', '首轮不通过，二轮修正'],
 ['接口隔离', '预约仓储接口、身份接口', 'IdentityInterface、ReservationRepository', '通过'],
 ['并发处理', '唯一约束兜底，未写异常转译链', '三轮给出查冲突、唯一约束、异常转译、友好提示的完整链路', 'AI 更完整'],
 ['关联语义', '预约申请与时段为关联，时段可被多个申请引用', '修正版用组合 Reservation *-- Period', '不通过：时段随申请销毁'],
 ['命名一致性', '中文类名，与分析模型同构', '第一轮中文，第二轮改为英文类名加中文注释', '部分通过'],
 ['跨图一致', '顺序图中的对象都能在类图找到', '顺序图的 Repository、Identity 有对应；Period 未出现', '部分通过'],
 ['幻觉', '—', '微服务、消息代理均属方案空间内的合理内容', '通过']], [2.6, 5.2, 5.6, 2.6])
add_p('九项中五项通过或部分通过，两项不通过（依赖方向、关联语义），两项需分轮次判断。值得注意的是第二轮：模型在提示词点名“职责分离”后，确实把校验从实体移到了服务，说明具体到可判定的约束能被纠正；而对“组合还是关联”没人提醒，它反而把原本正确的关联改成了组合。')

heading('12.3 方案论证', 2)
table(['维度', '方案A 分层单体', '方案B 事件驱动'], [
 ['可维护性', '单部署单元、本地事务，日志与排障集中；改动影响面可估', '需维护消息代理、多服务与死信队列；跨服务问题定位成本高'],
 ['可扩展性', '整体复制扩容，粒度粗', '按服务独立扩容，粒度细'],
 ['需求匹配', '单服务器、3 秒响应、审核结果即时可见，均满足', '异步化使审核结果延迟，最终一致与本系统的即时可见要求冲突']], [2.6, 6.2, 6.2])
add_p('三个维度上，方案A 在可维护性与需求匹配度上明显占优，方案B 只在可扩展性上领先；而本系统的规模与部署条件并不需要那种扩展粒度。因此选型结论是采用与手工设计一致的分层单体，并从 AI 输出中采纳两点改进：其一，把“查冲突 + 唯一约束 + 异常转译 + 友好提示”写成完整的并发处理链，此前 5.3 节只写到唯一约束兜底；其二，明确实体不承担业务编排，与手工设计中把审核行为上移到服务层的取向一致。')

add_p('关于活动 4 要求的“小组内以 Git 提交历史与 diff 进行同伴互评”：本实验为个人独立完成，无组内互评环节。作为替代，本人在提交历史中逐次记录设计决策，并以 12.2 的逐项对照代替他人评审；若后续组队，可直接以本仓库的 diff 作为互评素材。')

heading('12.4 缺陷清单与改进建议', 2)
table(['编号', '所属图', '缺陷', '类型', '改进建议'], [
 ['D1', '图 8', '预约申请指向预约仓储实现，且把 interface 构造型标在实现上', '依赖方向', '改为实体不依赖仓储；仓储接口由服务依赖，实现放基础设施层'],
 ['D2', '图 9', '预约与审核各自独立数据库，未交代跨库查询与一致性', '一致性', '若采用，需补事件契约与补偿策略，否则维持单库'],
 ['D3', '图 10', '容量校验、冲突判断放在可预约时段实体', '职责分配', '校验编排上移到服务，实体只做自身不变式（如时段起止合法）'],
 ['D4', '图 12', '预约申请与时段用组合，时段随申请销毁', '关联语义', '改为关联，时段独立存在并被多个申请引用'],
 ['D5', '图 10、12', '中文注释与英文类名混用', '命名', '统一为中文类名，与实验一、二的分析模型同构'],
 ['D6', '图 11', 'autonumber 与手写序号叠加，渲染为“1 1.”', '图面', '去掉手写序号或关闭 autonumber']], [1.4, 1.8, 6.6, 1.8, 4.4])
add_p('改进建议分两层。模型层：D1 与 D4 影响结构正确性，必须修正；D3 已在第二轮自愈，可保留修正版。提示词层：本轮把“职责分离”写进提示词后模型确实照做，说明可判定的约束有效；但“关联还是组合”没有被点名，模型就把原本正确的关联改错了。下一轮应把易错点逐条列成检查项，而不是依赖模型自觉。')

href = heading('参考文献', 1); href.paragraph_format.page_break_before = True
add_p('[1] 王健雄, 冯元勇. 面向对象分析与设计实验指导书 V202610[Z]. 广州大学计算机科学与网络工程学院, 2026.', indent=False)
add_p('[2] Object Management Group. OMG Unified Modeling Language Version 2.5.1[EB/OL]. https://www.omg.org/spec/UML/2.5.1, 2017[2026-11-16].', indent=False)
add_p('[3] PlantUML. PlantUML Language Reference Guide[EB/OL]. https://plantuml.com/, 2026[2026-11-16].', indent=False)

# ---------------- 附录 ----------------
happ = heading('附录 A  PlantUML 源文件清单', 1)
happ.paragraph_format.page_break_before = True
add_p('本实验的七张模型图均由 diagrams/ 下的源文件渲染，源文件与报告一并提交。')
table(['源文件', '图号', '内容'], [
 ['01_分层体系结构.puml', '图 1', '四层职责与依赖方向'],
 ['02_设计类图.puml', '图 3', '服务、实体、接口与状态枚举'],
 ['03_部署图.puml', '图 2', '浏览器、应用服务器、数据库与身份服务'],
 ['04_提交预约顺序图.puml', '图 5', '设计级消息序列与事务边界'],
 ['05_预约申请状态图.puml', '图 6', '状态集合与守卫条件'],
 ['06_界面原型.puml', '图 7', '提交预约界面线框'],
 ['07_数据模型.puml', '图 4', '五张表与主外键关系']], [7.4, 1.8, 6.4])

happ2 = heading('附录 B  AI 原始输出', 1)
add_p('以下为两轮提示词下模型返回的 PlantUML 原文，未作修改，与 12.1 的五张图一一对应。')

for label, fname in [('B.1 方案A 分层单体（对应图 8）', 'ai_A_分层单体.puml'),
                     ('B.2 方案B 事件驱动（对应图 9）', 'ai_B_事件驱动.puml'),
                     ('B.3 核心类图第一轮（对应图 10）', 'ai_类图_第一轮.puml'),
                     ('B.4 设计级顺序图（对应图 11）', 'ai_顺序图_并发.puml'),
                     ('B.5 核心类图修正版（对应图 12）', 'ai_类图_第二轮.puml')]:
    add_p(label, indent=False)
    for line in (DIAG / fname).read_text(encoding='utf-8').splitlines():
        cp = doc.add_paragraph()
        cp.paragraph_format.first_line_indent = Pt(0)
        cp.paragraph_format.line_spacing = 1.0
        cp.paragraph_format.space_after = Pt(0)
        cr = cp.add_run(line if line else ' ')
        cr.font.name = 'Consolas'
        cr._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), '宋体')
        cr.font.size = Pt(8)
    add_p('')

# ---------------- 收尾 ----------------
while doc.paragraphs and not doc.paragraphs[-1].text.strip() and not doc.paragraphs[-1].runs:
    last = doc.paragraphs[-1]
    last._p.getparent().remove(last._p)

for p in doc.paragraphs:
    for r in p.runs:
        if r.font.size is None:
            font(r)

doc.core_properties.title = '校园实验室预约管理系统系统设计实验报告'
doc.core_properties.author = '封媛'
doc.core_properties.subject = '面向对象分析与设计实验二'
doc.save(OUT)
print(OUT)
