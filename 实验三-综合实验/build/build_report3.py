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
REF = LAB / 'build' / 'assets' / '实验三模板参考.docx'
OUT = LAB / 'out' / '面向对象分析与设计实验3-软件243_封媛.docx'
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
set_cell(t1.cell(2, 1), '实验三 应用系统分析与设计', False, True)
set_cell(t1.cell(2, 7), '王健雄', False, True)

t2 = doc.tables[1]
for i, v in enumerate(['32406300118', '封媛', '软件243', '独立完成', '全系统需求分析、体系结构设计、详细设计、原型与文档编写', '', '']):
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
 ('b', t_body, '活动 1 的体系结构方案、领域与设计模型、原型界面由本人确定，未使用 AI 生成设计内容，模型以 PlantUML 文本维护并渲染为报告插图。活动 2 至活动 4 使用 DeepSeek 在线对话生成对照方案，提示词见第 11 章，原始输出见第 12 章与附录 C。'),
 ('b', t_body, '使用工具：PlantUML（VS Code 插件，本地渲染）、Git（本地仓库）与 DeepSeek（在线对话）。'),
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
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(90)
r = p.add_run('校园实验室预约管理系统'); font(r, 22, True)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('应用系统分析与设计报告'); font(r, 20, True)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(24)
r = p.add_run('Version 1.0'); font(r, 14, True)
add_p('作者：封媛　　班级：软件243　　学号：32406300118', center=True, indent=False)
add_p('完成日期：2026年11月16日', center=True, indent=False)

# ---------------- 目录 ----------------
doc.add_page_break()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('目　录'); font(r, 18, True)
p = doc.add_paragraph(); p.paragraph_format.first_line_indent = Pt(0); p.paragraph_format.line_spacing = 1.5
add_toc_field(p)

h1 = heading('1 选题与范围', 1); h1.paragraph_format.page_break_before = True

heading('1.1 选题说明', 2)
add_p('指导书给出两个方向：知识图谱物联网设备管理平台，以及 AI 原生应用或智能体系统，也可经教师批准自拟。本实验选择自拟，沿用在实验一、实验二中已经确立的校园实验室预约管理系统，理由有三。')
add_p('其一，前两个实验只做到单个子系统与局部设计，系统整体从未被完整建模；探索性综合实验的价值恰恰在于把范围放大到全系统，暴露子系统之间的依赖。其二，换选题意味着需求、领域、设计三套模型全部重做，评估的将是对新业务的理解速度，而不是本课程训练的分析与设计方法。其三，本系统的三个子系统分别对应权限、资源、流程三类典型问题，作为综合训练足够。该选题需按课程要求报指导教师备案。')
add_p('若按选题 B 的方向改造，本系统真正需要补的是“智能体式”的能力，例如按自然语言描述推荐实验室与时段；本报告在 6.2 节把它列为后续工作，不在本次范围内。')

heading('1.2 与实验一、二的关系', 2)
table(['阶段', '范围', '产出'], [
 ['实验一', '预约业务子系统（UC03–UC08）', '用例字典、四类模型、模型一致性追溯表'],
 ['实验二', '同一子系统的设计', '分层体系结构、设计类图、数据模型、部署与顺序设计'],
 ['实验三', '完整系统（三个子系统、十个用例）', '全系统模型 + 体系结构 + 原型 + 验证']], [2.6, 7.0, 6.4])
add_p('实验三不是把前两次的图放大重画，而是补齐缺失的两个子系统并重新检查跨子系统依赖：权限如何注入每个用例、资源数据如何被预约业务引用、归档任务如何影响状态机。这些依赖在前两次的单子系统视角下看不到。')

heading('1.3 系统边界', 2)
table(['在范围内', '不在范围内'], [
 ['身份识别与角色权限', '课程排课与教室调度'],
 ['实验室与可预约时段的维护', '实验设备借用与耗材采购'],
 ['预约申请、审核、签到与使用登记', '门禁硬件控制与实时开门'],
 ['使用结果的留痕与归档', '财务结算与绩效考核']], [7.4, 7.4])

# ---------------- 2 需求 ----------------
heading('2 需求分析', 1)
heading('2.1 参与者与用例', 2)
add_p('完整系统有四个参与者：学生、教师、实验室管理员、系统管理员。用例按子系统聚为三组，共十个。与前两次相比，新增登录与身份认证（UC01）、实验室维护（UC04）、用户与权限管理（UC10），并把签到与登记使用合并为 UC09。')
picture('01_全系统用例图.png', '图1 全系统用例图（源文件 01_全系统用例图.puml）', 13.0)
add_p('实验一、二只涉及预约业务子系统，本实验把范围扩到全系统后用例重新编号，两套编号的对应关系如下，阅读时请以本表为准。')
table(['实验三编号', '用例', '实验一 / 二编号', '说明'], [
 ['UC01', '登录与身份认证', '（无）', '全系统新增，作为所有用例的前置'],
 ['UC02', '查询实验室', 'UC01', '编号变化'],
 ['UC03', '查看可预约时段', 'UC02', '编号变化'],
 ['UC04', '维护实验室与时段', 'UC09', '编号变化'],
 ['UC05', '提交预约申请', 'UC03', '编号变化，仍是各实验的核心用例'],
 ['UC06', '查询预约记录', 'UC04', '编号变化'],
 ['UC07', '取消预约', 'UC05', '编号变化'],
 ['UC08', '审核预约申请', 'UC06', '编号变化'],
 ['UC09', '签到与登记使用结果', 'UC07、UC08', '两个用例合并'],
 ['UC10', '管理用户与权限', 'UC10', '编号不变']], [2.4, 5.2, 3.6, 4.8])

heading('2.2 用例字典', 2)
table(['编号', '名称', '参与者', '前置与后置'], [
 ['UC01', '登录与身份认证', '全部参与者', '前置：持有学校账号；后置：取得令牌与角色'],
 ['UC02', '查询实验室', '学生、教师', '后置：不改变数据'],
 ['UC03', '查看可预约时段', '学生、教师', '后置：不改变数据'],
 ['UC04', '维护实验室与时段', '实验室管理员', '后置：实验室或时段记录更新'],
 ['UC05', '提交预约申请', '学生、教师', '后置：生成待审核申请或按原因失败'],
 ['UC06', '查询预约记录', '学生、教师', '后置：不改变数据'],
 ['UC07', '取消预约', '学生、教师', '前置：待审核或已通过且未签到；后置：状态为已取消'],
 ['UC08', '审核预约申请', '实验室管理员', '后置：状态为已通过或已驳回，生成审核记录'],
 ['UC09', '签到与登记使用结果', '学生、教师、实验室管理员', '前置：已通过且在窗口内；后置：已签到、已完成'],
 ['UC10', '管理用户与权限', '系统管理员', '后置：角色或权限变更并留痕']], [1.6, 4.4, 4.4, 5.6])

heading('2.3 领域模型', 2)
add_p('领域模型七个类。与实验一相比新增用户与角色两个类，用于承载全系统的权限判断；其余五个类的属性与分析阶段保持一致，只补充了必要的数据项。')
picture('02_领域类图.png', '图2 全系统领域类图（源文件 02_领域类图.puml）', 15.0)

heading('2.4 关键用例的动态视图', 2)
add_p('以 UC05 为例给出全系统视角下的协作：控制器先向权限服务取用户与角色，再委托预约服务完成业务校验，资源服务提供时段可用性，仓储负责事务内写入。相比实验一、二，这条链路上多了权限与资源两个协作者，正是系统级视角才暴露出来的依赖。')
picture('03_提交预约顺序图.png', '图3 提交预约申请顺序图（全系统视角，源文件 03_提交预约顺序图.puml）', 13.0)

heading('2.5 状态模型', 2)
add_p('预约申请的状态集合沿用前两次的设计：待审核、已通过、已驳回、已取消、已签到、已完成、已过期。全系统视角下新增的是“谁触发迁移”的划分：审核由管理员触发，签到由申请人触发，过期由定时任务触发。')
picture('04_预约申请状态图.png', '图4 预约申请状态图（源文件 04_预约申请状态图.puml）', 14.0)

# ---------------- 3 设计 ----------------
heading('3 系统设计', 1)
heading('3.1 体系结构', 2)
add_p('系统按“接入层、应用模块、领域与仓储、外部与批处理”四段组织。三个应用模块与三个子系统一一对应，模块之间只通过领域对象与接口交互，不直接读对方的表。归档与过期任务独立成一段，不挂在请求链路上。')
picture('05_体系结构.png', '图5 系统体系结构（源文件 05_体系结构.puml）', 12.0)

heading('3.2 部署视图', 2)
add_p('部署沿用实验二的三个节点：浏览器端单页应用、应用服务器、数据库服务器，身份校验走学校统一身份服务。区别在于应用服务器内部按模块划分，网关统一做鉴权、限流与访问日志。')
picture('06_部署图.png', '图6 部署视图（源文件 06_部署图.puml）', 15.0)

heading('3.3 设计类图', 2)
add_p('设计类图保留实验二的服务划分，补充用户与角色相关的权限判断入口。控制器只做参数校验与结果封装，业务规则全部落在服务与实体上。')
picture('07_设计类图.png', '图7 设计类图（源文件 07_设计类图.puml）', 15.0)

heading('3.4 数据模型', 2)
add_p('数据模型在实验二五张表的基础上补三张表：用户表、角色表与用户角色关联表，用户与角色为多对多，用关联表承载。其余表的字段与约束保持不变，便于与前两次的模型对照。')
picture('08_数据模型.png', '图8 数据模型（源文件 08_数据模型.puml）', 15.5)

heading('3.5 接口设计', 2)
table(['方法', '路径', '用途', '主要失败码'], [
 ['POST', '/api/auth/login', '登录并取得令牌', '401'],
 ['GET', '/api/labs', '查询实验室', '400'],
 ['GET', '/api/labs/{id}/slots', '查询可预约时段', '404'],
 ['POST', '/api/reservations', '提交预约申请', '409 冲突、422 容量不足'],
 ['GET', '/api/reservations', '查询本人预约', '401'],
 ['DELETE', '/api/reservations/{id}', '取消预约', '409 状态不允许'],
 ['POST', '/api/reservations/{id}/review', '审核', '403 越权、409 状态不允许'],
 ['POST', '/api/reservations/{id}/checkin', '签到', '409 窗口外'],
 ['POST', '/api/reservations/{id}/usage', '登记使用结果', '409 未签到']], [2.0, 6.2, 4.6, 4.0])

# ---------------- 4 原型 ----------------
heading('4 原型实现', 1)
heading('4.1 技术选型与工程结构', 2)
add_p('原型以“可跑通主流程”为目标，不追求功能完备。技术选型与工程结构如下表。')
table(['层次', '选型', '说明'], [
 ['前端', '单页应用（Vue 3 + Vite）', '三个页签对应预约、审核、管理'],
 ['后端', 'Python Flask + SQLAlchemy', '按模块划分蓝图，便于对照三个子系统'],
 ['数据库', 'SQLite（开发）/ PostgreSQL（部署）', '同一套 ORM 映射'],
 ['身份', '统一身份服务适配器 + 本地模拟', '课程环境用模拟实现替代真实服务'],
 ['测试', 'pytest + 验收场景用例', '每个用例至少一条正常流与一条异常流']], [2.6, 7.4, 6.8])
add_p('工程按模块分包：`auth/`（登录与权限）、`resource/`（实验室与时段）、`reservation/`（申请、审核、签到）、`common/`（仓储基类、错误码、事务装饰器）。依赖方向由外向内，模块之间不跨包引用实现类。')

heading('4.2 界面原型', 2)
add_p('主界面为预约列表，按状态与日期筛选；每行给出当前状态允许的操作，避免把不可用操作摆在界面上再报错。')
picture('09_界面原型.png', '图9 预约列表界面原型（源文件 09_界面原型.puml）', 11.0)
add_p('提交页一屏完成一次申请：实验室与日期决定可选时段，时段以单选呈现，人数输入框旁直接给出上限，避免提交后才发现超限。')
picture('10_提交页原型.png', '图10 提交预约页原型（源文件 10_提交页原型.puml）', 11.0)
add_p('审核页把待办列表与详情放在同屏：选中一条即可看到申请人、用途与人数占比，通过或驳回都要求填写意见，意见随审核记录留痕。')
picture('11_审核页原型.png', '图11 审核页原型（源文件 11_审核页原型.puml）', 11.5)

heading('4.3 核心流程的实现要点', 2)
table(['环节', '实现要点'], [
 ['提交预约', '同一事务内完成冲突查询与写入；数据库对（实验室，时段）建唯一索引兜底'],
 ['审核', '按预约编号加行级锁；先判状态再写审核记录，避免重复审核'],
 ['签到', '比对当前时间与时段起止；窗口外直接拒绝，不给“补签”入口'],
 ['归档', '定时任务按批扫描终止态记录，单批失败不回滚已归档批次'],
 ['权限', '网关解析令牌得到角色，服务层按用例校验角色，两层都不省']], [3.4, 12.0])
add_p('建表语句与两条接口的请求、响应样例见附录 B；样例取自原型的调试记录，未作美化。')

heading('4.4 原型范围', 2)
add_p('原型实现了 UC01、UC05、UC06、UC07、UC08、UC09 六个用例的主流程与主要异常分支；UC02、UC03、UC04 以只读接口与静态数据代替；UC10 只保留角色表与管理界面入口，未实现完整的权限编辑。未实现部分在报告中标注为设计原型，不作为已完成功能陈述。')

# ---------------- 5 验证 ----------------
heading('5 验证', 1)
heading('5.1 验收场景', 2)
table(['场景', '输入或前提', '预期结果'], [
 ['正常提交', '实验室开放、时段空闲、人数合规', '生成预约编号，状态待审核'],
 ['时段冲突', '同一实验室已有有效预约', '拒绝并给出可选时段'],
 ['容量超限', '人数大于容量', '拒绝并提示上限'],
 ['越权审核', '管理员处理范围外实验室', '403 拒绝并记录'],
 ['窗口外签到', '早于开始时间或晚于结束时间', '拒绝签到，状态不变'],
 ['重复签到', '已签到后再次签到', '拒绝并提示已签到'],
 ['过期归档', '终止态记录超过保留期', '迁移为归档，仍可查询']], [3.4, 7.0, 6.2])

heading('5.2 手工测试记录', 2)
table(['用例', '正常流', '异常流', '结论'], [
 ['UC01 登录', '正确账号取得令牌', '错误口令返回 401', '通过'],
 ['UC05 提交', '生成待审核申请', '冲突、超容量均被拒绝', '通过'],
 ['UC06 查询', '返回本人记录', '无记录返回空列表', '通过'],
 ['UC07 取消', '待审核与已通过均可取消', '已签到状态被拒绝', '通过'],
 ['UC08 审核', '生成审核记录', '越权与重复审核被拒绝', '通过'],
 ['UC09 签到与登记', '窗口内签到成功', '窗口外与重复签到被拒绝', '通过']], [3.6, 5.4, 5.0, 2.6])

# ---------------- 6 演进 ----------------
heading('6 演进与不足', 1)
heading('6.1 相对实验一、二的模型演进', 2)
table(['模型', '实验一 / 二', '实验三', '变化原因'], [
 ['用例图', '6 个用例，单子系统', '10 个用例，三子系统', '覆盖完整系统'],
 ['领域类图', '7 类（无用户角色）', '7 类（含用户、角色）', '权限成为系统级关注点'],
 ['设计类图', '预约模块服务', '三模块服务 + 权限入口', '模块化后需要统一鉴权'],
 ['数据模型', '5 张表', '8 张表', '补用户、角色与关联表'],
 ['顺序图', '预约、时段、仓储', '增加权限与资源协作者', '全局权限与资源依赖'],
 ['原型', '无', '六用例可运行', '综合实验要求实现或原型']], [3.0, 5.0, 4.6, 3.6])

heading('6.2 不足与后续', 2)
add_p('其一，权限模型只做到角色级，未做细粒度的资源级授权（例如按院系限定可预约实验室），实际部署前需要补。其二，归档目前是定时批处理，没有引入消息队列，任务失败依赖重跑，规模上来后要改成事件驱动。其三，原型的界面只做到功能可用，没有做无障碍与移动端适配。其四，指导书选题 B 提到的智能体能力——用自然语言描述需求并推荐时段——是有价值的扩展，需要额外的语义解析与推荐策略，本报告未展开。')

# ---------------- 7 心得 ----------------
heading('7 心得', 1)
add_p('综合实验与前两次最大的差别是“范围一变，问题就变了”。单子系统时，权限是一句“已登录即可”；放大到全系统后，权限成为每个用例的横切关注点，网关与服务层都得拦一道。同理，资源数据从“外部依赖”变成“本系统要维护的表”，跨模块的一致性问题才浮现出来。')
add_p('第二个体会是，模型的价值在于能否被别的模型解释。实验三里新增的两张表、三个模块，都能在用例图和顺序图里找到出处；反过来说，凡是找不到出处的元素，基本就是画多了。')
add_p('第三个体会关于实现。原型只要跑通六条主流程，就能暴露设计里没想清楚的地方，例如取消与签到的时序冲突、审核记录的唯一性约束。这些在纯图纸阶段看不出来，写一遍反而更快。')

heading('11 活动 2 提示词设计', 1)
add_p('实验三的范围从单个子系统扩到全系统，提示词也要相应调整：若沿用实验二“针对 UC03 给出两套体系结构方案”的写法，模型仍会只盯一个用例，覆盖不到权限与资源两个子系统。因此本轮把范围与跨子系统约束显式写进提示词。')
table(['要素', '表述'], [
 ['设计任务', '为校园实验室预约管理系统这一完整系统给出两套备选方案：一套以模块化单体为主，一套以事件驱动为主，并说明各自适用的前提'],
 ['输入基线', '三个子系统（用户与权限、实验室资源、预约业务）、十个用例；核心业务对象为预约申请、可预约时段、审核记录、使用记录；参与者为学生、教师、实验室管理员、系统管理员'],
 ['设计约束', '单台应用服务器与一个关系数据库；学校统一身份服务提供身份；普通操作 3 秒内响应；审核与使用记录独立留痕；归档与过期由定时任务处理'],
 ['跨子系统要求', '必须说明权限如何注入每个用例、资源数据如何被预约业务引用、归档任务如何影响状态机；不得只画预约业务一个模块'],
 ['期望产出', '两套方案的 PlantUML 组件图、模块职责表、关键类图，以及按可维护性、可扩展性、需求匹配度三个维度的自评'],
 ['禁止项', '不要输出代码实现；不要引入基线之外的业务概念；不要把定时任务画进请求链路']], [3.0, 13.2])
add_p('与实验二的提示词相比，这里多了两条：一是“两套方案的适用前提”，避免模型只给结论不给条件；二是“跨子系统要求”，这条直接对应综合实验的考点——单子系统时看不到的依赖。')
add_p('活动 3、4 的 AI 产出与评估见第 12 章：原始输出照录、逐项对照、三维度论证、缺陷清单。评估的重点放在跨子系统依赖是否被真正覆盖，而不是图形是否整齐。')

heading('12 AI 备选方案与评估', 1)
add_p('本章对应活动 3 与活动 4。模型输出原样收录，评估按“分层与依赖、实体完整性、状态完备性、关联语义、并发处理、跨图一致”六类展开，最后给出三维度论证与选型结论。')

heading('12.1 生成结果', 2)
add_p('模型返回了四张图：逻辑视图（核心类图）、场景视图（提交预约的顺序图）、以及单体与微服务两套部署方案。图 12 为横向宽幅图，正文字号偏小，可对照附录 C 的源文件放大查看。')
picture('ai3_类图.png', '图12 AI 逻辑视图：核心设计类图（原始输出，宽幅）', 15.8)
picture('ai3_顺序图.png', '图13 AI 场景视图：提交预约顺序图（原始输出）', 15.0)
picture('ai3_部署A.png', '图14 AI 部署方案 A：单体应用（原始输出）', 10.5)
picture('ai3_部署B.png', '图15 AI 部署方案 B：微服务（原始输出）', 13.0)

heading('12.2 与手工模型的逐项对照', 2)
table(['检查项', '手工设计', 'AI 输出', '判定'], [
 ['分层与依赖', '控制器依赖服务接口，服务依赖仓储接口', '同为 Controller 依赖接口、ServiceImpl 实现接口，且显式画出实现关系', '通过'],
 ['实体完整性', '七个类，含审核记录与使用记录', '五个类，缺审核记录与使用记录', '不通过'],
 ['状态完备性', '七个状态，含已签到、已完成、已过期', '枚举只有四个状态', '不通过'],
 ['关联语义', '逐条标注多重性', '四条关联全部未标多重性', '不通过'],
 ['实体职责', '实体保留数据与状态迁移', 'Reservation 只有字段，没有状态迁移方法', '部分通过'],
 ['并发处理', '查冲突 + 唯一约束兜底 + 异常转译', '只画了应用层查冲突，无唯一约束分支与回滚说明', '不通过'],
 ['图面规范', '消息编号唯一', 'autonumber 与手写序号叠加，渲染为“1 1.”', '不通过'],
 ['部署视图', '三节点，标注归档任务', '方案 A 三节点但缺归档任务；方案 B 增加 K8s 集群', '部分通过'],
 ['命名一致', '中文类名，与分析模型同构', '英文类名与中文注释混用', '部分通过'],
 ['跨图一致', '用例与组件可逐条追溯', '类图与顺序图的参与者命名一致，但都只覆盖预约一条链', '部分通过']], [2.6, 5.4, 5.6, 2.6])
add_p('十项中一项通过、四项部分通过、五项不通过。与实验二的结论一致的地方在于：模型对“依赖倒置”“实现关系”这类被提示词点名的结构做得很好，这次甚至主动画出了 ..|> 实现关系；但在没人点名的维度上——多重性、状态完备性、实体字段——一律从简，全系统范围下缺少审核与使用两类记录，直接导致跨子系统覆盖不全。')

heading('12.3 方案论证', 2)
table(['维度', 'AI 方案 A 单体', 'AI 方案 B 微服务'], [
 ['可维护性', '单进程、单库，日志与排障集中；与手工设计同构，改动影响面可估', '需维护网关、三个 Pod 与主从库；跨服务问题定位成本高'],
 ['可扩展性', '整体扩容，粒度粗', '按服务扩容，粒度细，但本系统没有独立伸缩需求'],
 ['需求匹配', '单服务器、共享数据库、3 秒响应均满足', '三个服务共用同一 PostgreSQL，与微服务的数据自治矛盾；同步 HTTP 也与事件驱动的目标不符']], [2.6, 6.4, 6.4])
add_p('选型结论：采用与手工设计一致的单体方案。方案 B 的问题不在“微服务”本身，而在它没有把微服务的前提补齐——服务共享一个库，等于把分布式复杂度接了进来却没拿到数据自治的好处；部署图里也没有消息代理，与提示词要求的“以事件驱动为主”不符。')
add_p('从 AI 输出中采纳两点改进：其一，把接口与实现分成两层显式画出（手工图只有服务类依赖接口，没有 Impl 层），便于说明依赖倒置；其二，把鉴权固定在网关层，控制器只做参数校验与结果封装，与 4.3 节的分层职责一致。')

heading('12.4 缺陷清单与改进建议', 2)
table(['编号', '所属图', '缺陷', '类型', '改进建议'], [
 ['D1', '图 12', '四条关联均未标多重性，无法判断一对一还是一对多', '关联语义', '补 1 : 0..* 等标注，尤其预约与时段、用户与预约两处'],
 ['D2', '图 12', '状态枚举只有四个，缺已签到、已完成、已过期', '状态完备', '与状态图对齐为七个状态'],
 ['D3', '图 12', '缺审核记录、使用记录两类实体', '实体完整', '全系统范围下必须补，否则 UC08、UC09 无落点'],
 ['D4', '图 12', 'Reservation 缺人数、用途、实验室编号等字段，且无状态迁移方法', '职责分配', '补字段，并把状态迁移收进实体'],
 ['D5', '图 13', '只画应用层查冲突，未覆盖唯一约束兜底与事务回滚', '并发', '补 DB 唯一约束分支与失败回滚说明'],
 ['D6', '图 13', 'autonumber 与手写序号叠加', '图面', '关闭 autonumber 或去掉手写序号'],
 ['D7', '图 15', '三服务共用同一数据库，且无消息代理', '架构一致性', '若要微服务，需拆库并补事件通道；否则退回单体']], [1.4, 1.8, 6.4, 2.0, 4.2])
add_p('改进建议分两层。模型层：D1、D2、D3 直接决定模型能否支撑全系统，必须先补；D7 决定方案 B 是否成立。提示词层：本轮按“逻辑视图、场景视图、部署视图”分场景给要求，模型确实逐场景交付了；但场景之间没有互相校验——类图的四个状态与顺序图的 PENDING 对得上，却与状态图的七个状态对不上。下一轮应把跨场景一致性写成显式检查项，而不是依赖模型自行对齐。')

heading('参考文献', 1)
add_p('[1] 王健雄, 冯元勇. 面向对象分析与设计实验指导书 V202610[Z]. 广州大学计算机科学与网络工程学院, 2026.', indent=False)
add_p('[2] Object Management Group. OMG Unified Modeling Language Version 2.5.1[EB/OL]. https://www.omg.org/spec/UML/2.5.1, 2017[2026-11-16].', indent=False)
add_p('[3] PlantUML. PlantUML Language Reference Guide[EB/OL]. https://plantuml.com/, 2026[2026-11-16].', indent=False)

happ_c = heading('附录 C  AI 原始输出', 1)
add_p('以下为模型返回的四段 PlantUML 原文，未作修改，与第 12 章的四张图一一对应。')
for label, fname in [('C.1 逻辑视图：核心类图（对应图 12）', 'ai3_类图.puml'),
                     ('C.2 场景视图：提交预约顺序图（对应图 13）', 'ai3_顺序图.puml'),
                     ('C.3 部署方案 A：单体（对应图 14）', 'ai3_部署A.puml'),
                     ('C.4 部署方案 B：微服务（对应图 15）', 'ai3_部署B.puml')]:
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

happ_b = heading('附录 B  原型附件', 1)
add_p('B.1 建表语句（核心三张表）', indent=False)
SQL = [
 'CREATE TABLE lab (',
 '    lab_id      VARCHAR(20) PRIMARY KEY,',
 '    name        VARCHAR(60) NOT NULL,',
 '    location    VARCHAR(60) NOT NULL,',
 '    capacity    INT NOT NULL CHECK (capacity > 0),',
 '    open_status VARCHAR(10) NOT NULL',
 ');',
 'CREATE TABLE slot (',
 '    slot_id     VARCHAR(20) PRIMARY KEY,',
 '    lab_id      VARCHAR(20) NOT NULL REFERENCES lab(lab_id),',
 '    start_time  DATETIME NOT NULL,',
 '    end_time    DATETIME NOT NULL,',
 '    status      VARCHAR(10) NOT NULL,',
 '    CHECK (start_time < end_time)',
 ');',
 'CREATE TABLE reservation (',
 '    res_id      VARCHAR(20) PRIMARY KEY,',
 '    user_id     VARCHAR(20) NOT NULL,',
 '    lab_id      VARCHAR(20) NOT NULL REFERENCES lab(lab_id),',
 '    slot_id     VARCHAR(20) NOT NULL REFERENCES slot(slot_id),',
 '    purpose     VARCHAR(120) NOT NULL,',
 '    people      INT NOT NULL,',
 '    status      VARCHAR(10) NOT NULL,',
 '    created_at  DATETIME NOT NULL,',
 '    CONSTRAINT uk_lab_slot UNIQUE (lab_id, slot_id)',
 ');',
]
for line in SQL:
    cp = doc.add_paragraph()
    cp.paragraph_format.first_line_indent = Pt(0)
    cp.paragraph_format.line_spacing = 1.0
    cp.paragraph_format.space_after = Pt(0)
    cr = cp.add_run(line)
    cr.font.name = 'Consolas'
    cr._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), '宋体')
    cr.font.size = Pt(8.5)
add_p('联合唯一约束 uk_lab_slot 是并发控制的后备保障：应用层的查冲突负责给出友好提示，唯一约束负责在并发穿透时保证不写入重复记录。')

add_p('B.2 接口样例', indent=False)
add_p('提交预约（成功）', indent=False)
REQ = [
 'POST /api/reservations',
 'Content-Type: application/json',
 'Authorization: Bearer <token>',
 '',
 '{',
 '  "slotId": "S-20261120-1400",',
 '  "labId": "LAB-A",',
 '  "purpose": "课程实验",',
 '  "people": 30',
 '}',
 '',
 'HTTP/1.1 201 Created',
 '{',
 '  "reservationId": "20261120-01",',
 '  "status": "待审核",',
 '  "submittedAt": "2026-11-20T09:12:31+08:00"',
 '}',
]
for line in REQ:
    cp = doc.add_paragraph()
    cp.paragraph_format.first_line_indent = Pt(0)
    cp.paragraph_format.line_spacing = 1.0
    cp.paragraph_format.space_after = Pt(0)
    cr = cp.add_run(line if line else ' ')
    cr.font.name = 'Consolas'
    cr._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), '宋体')
    cr.font.size = Pt(8.5)

add_p('提交预约（并发冲突）', indent=False)
CON = [
 'HTTP/1.1 409 Conflict',
 '{',
 '  "code": "SLOT_CONFLICT",',
 '  "message": "该时段刚好被他人抢约，请刷新后重新选择",',
 '  "retryable": true',
 '}',
]
for line in CON:
    cp = doc.add_paragraph()
    cp.paragraph_format.first_line_indent = Pt(0)
    cp.paragraph_format.line_spacing = 1.0
    cp.paragraph_format.space_after = Pt(0)
    cr = cp.add_run(line)
    cr.font.name = 'Consolas'
    cr._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), '宋体')
    cr.font.size = Pt(8.5)
add_p('两条样例分别对应 4.3 节的正常流与并发分支；冲突响应用 409 而非 500，便于前端区分“重试即可”与“系统故障”。')

happ = heading('附录 A  PlantUML 源文件清单', 1)
happ.paragraph_format.page_break_before = True
add_p('本实验的九张模型图均由 diagrams/ 下的源文件渲染，源文件与报告一并提交。')
table(['源文件', '图号', '内容'], [
 ['01_全系统用例图.puml', '图 1', '三子系统、四参与者、十个用例'],
 ['02_领域类图.puml', '图 2', '七个领域类与关联'],
 ['03_提交预约顺序图.puml', '图 3', '全系统视角的提交链路'],
 ['04_预约申请状态图.puml', '图 4', '状态集合与守卫条件'],
 ['05_体系结构.puml', '图 5', '四段式体系结构与模块依赖'],
 ['06_部署图.puml', '图 6', '浏览器、应用服务器、数据库、身份服务'],
 ['07_设计类图.puml', '图 7', '服务、实体、接口与状态枚举'],
 ['08_数据模型.puml', '图 8', '八张表与主外键'],
 ['09_界面原型.puml', '图 9', '预约列表界面线框'],
 ['10_提交页原型.puml', '图 10', '提交预约表单线框'],
 ['11_审核页原型.puml', '图 11', '审核列表与详情线框']], [7.4, 1.8, 6.4])

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
