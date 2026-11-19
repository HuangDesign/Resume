# -*- coding: utf-8 -*-
"""黄红艺 · 产品经理简历（A4 打印版，编辑式排版 + STAR 项目）。"""

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

SERIF = "Didot"          # 拉丁衬线
SERIF_CN = "Songti SC"   # 中文宋体
SANS = "HarmonyOS Sans SC"

INK = RGBColor(0x2E, 0x2E, 0x2E)
BODY = RGBColor(0x4D, 0x4D, 0x4D)
MUTED = RGBColor(0x99, 0x99, 0x99)
RULE = "DBDBDB"

CONTENT_W = Cm(18.0)


# ---------- 基础工具 ----------
def set_font(run, size, *, serif=False, bold=False, color=BODY, spacing=None):
    lat = SERIF if serif else SANS
    cjk = SERIF_CN if serif else SANS
    run.font.name = lat
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:ascii"), lat)
    rf.set(qn("w:hAnsi"), lat)
    rf.set(qn("w:eastAsia"), cjk)
    if spacing is not None:
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(int(spacing * 20)))
        rpr.append(sp)


def para(doc, before=0, after=0, line=1.5, align=None,
         left=0, hanging=0, keep_next=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.widow_control = True
    if align is not None:
        p.alignment = align
    if left:
        pf.left_indent = Cm(left)
    if hanging:
        pf.first_line_indent = Cm(-hanging)
    if keep_next:
        pf.keep_with_next = True
    return p


def rule(p, size=4, color=RULE, space=6):
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color)
    bdr.append(bottom)
    pPr.append(bdr)


def rich(p, text, size, *, color=BODY, strong=INK, serif=False):
    """**加粗** 表示强调（改用深色而非加粗，保持版式轻）。"""
    for i, chunk in enumerate(text.split("**")):
        if not chunk:
            continue
        r = p.add_run(chunk)
        set_font(r, size, serif=serif, bold=False, color=(strong if i % 2 else color))


def section(doc, num, title, before=4):
    p = para(doc, before=before, after=5, line=1.0, keep_next=True)
    n = p.add_run(num + "  ")
    set_font(n, 7.2, color=MUTED, spacing=1.2)
    t = p.add_run(title)
    set_font(t, 12, serif=True, color=INK, spacing=0.6)
    rule(p, size=4, space=5)
    return p


# ---------- 文档 ----------
doc = Document()
normal = doc.styles["Normal"]
normal.font.name = SANS
normal.font.size = Pt(9)
normal.font.color.rgb = BODY
normal.element.rPr.rFonts.set(qn("w:eastAsia"), SANS)
normal.paragraph_format.space_before = Pt(0)
normal.paragraph_format.space_after = Pt(0)

sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(1.5)
sec.top_margin = Cm(1.15)
sec.bottom_margin = Cm(1.05)
sec.footer_distance = Cm(0.6)

# ---------- 页眉 ----------
head = doc.add_table(rows=1, cols=2)
head.autofit = False
head.columns[0].width = Cm(11.5)
head.columns[1].width = Cm(6.5)
left, right = head.rows[0].cells
for c, w in ((left, Cm(11.5)), (right, Cm(6.5))):
    c.width = w
    c._tc.get_or_add_tcPr().append(OxmlElement("w:tcBorders"))

p = left.paragraphs[0]
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
logo_run = p.add_run()
logo_run.add_picture("assets/logo.png", height=Cm(0.82))
gap = p.add_run("  ")
set_font(gap, 10)
nm = p.add_run("黄红艺")
set_font(nm, 21, serif=True, color=INK, spacing=0.6)

p2 = left.add_paragraph()
p2.paragraph_format.space_before = Pt(4)
p2.paragraph_format.space_after = Pt(0)
p2.paragraph_format.line_spacing = 1.0
rl = p2.add_run("产品经理 · AI 应用 / Agent / B 端平台")
set_font(rl, 9.2, color=BODY, spacing=0.2)

rp = right.paragraphs[0]
rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
rp.paragraph_format.space_after = Pt(0)
rp.paragraph_format.line_spacing = 1.45
for line in ("151 5889 3193  ·  942852369@qq.com",
             "GitHub github.com/HuangDesign",
             "公众号 uxd-si  ·  上海"):
    r = rp.add_run(line + "\n")
    set_font(r, 7.8, color=MUTED, spacing=0.2)

p = para(doc, before=5, after=0, line=1.0)
rule(p, size=8, space=0)

# ---------- 简介 ----------
p = para(doc, before=8, after=0, line=1.5)
rich(p, "多年企业级产品经验，深耕 IoT、协同办公、数据处理等业务的数智化发展领域。", 9.0, color=BODY)

p = para(doc, before=3, after=0, line=1.5)
rich(p,
     "近三年专注于 **AI Agent** 及**数据本体构建**，以透明可信、反馈进化、情境适配、责任可溯为产品理念，"
     "构建「1+1>2」的协作关系——让 AI 延伸人的能力，而非替代人。",
     9.6, color=INK)

# ---------- 能力维度 ----------
section(doc, "01", "能力维度")

DIMS = [
    ("业务判断", "先判断值不值得做，再谈怎么做——用目标、成本、可验证性筛选项。方案产出 **3 天 → 3 分钟**。"),
    ("用户洞察", "去现场看人怎么干活，而不是只看需求文档。试点 **12 位售前，10 位持续使用**。"),
    ("结构与交互", "角色、权限、数据流、本体先咬合清楚，页面才会收敛。**45 个组件规范**，入口 **3 → 1**。"),
    ("AI 与 Agent", "定义机器跑到哪一步必须停下等人确认，状态与回溯写进产品。覆盖 **近 1000 人**，满意度 **4.2**。"),
    ("数据与度量", "用行为数据、可用性测试与反馈闭环定位问题，把改动和结果对上账。满意度 **78% → 89%**。"),
    ("推动与交付", "与算法、合规、研发对齐边界与排期，在灰度里迭代。**7 所院校**部署，项目获公司优胜奖。"),
]
for label, text in DIMS:
    p = para(doc, before=1.6, after=0, line=1.32, left=2.4, hanging=2.4)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(2.4))
    lb = p.add_run(label + "\t")
    set_font(lb, 7.4, color=MUTED, spacing=0.6)
    rich(p, text, 8.7)

# ---------- 工作经历 ----------
section(doc, "02", "工作经历")

JOBS = [
    ("锐捷网络", "产品设计师（数据与 Agent 平台）", "2026.04 — 2026.10", [
        "负责 Source ID 平台**本体、数据管道、数据管理**三大模块的体验设计，并参与 Agent 任务链路与设计规范建设。",
        "把「数据接入—加工—建模—管理」收成一条统一路径，同时定义 Agent 过程状态、异常反馈与人工介入节点。",
        "用 Lovable 快速产出可交互原型验证方案，把任务流程、状态表达与反馈机制沉淀成团队原则。",
    ]),
    ("华为", "体验设计师 · Answer Agent 产品经理", "2022.04 — 2026.04", [
        "主导「智慧办公 Answer Agent」：组建团队、定产品定位与排期，完成从需求定义、Agent 工作流到灰度落地的全流程。",
        "负责 Talent 实训云平台的规划与落地，围绕教务、教学、考评梳理多角色任务与关键流程。",
    ]),
    ("涂鸦智能", "产品体验设计师（智慧楼宇 SaaS）", "2021.04 — 2022.04", [
        "负责 3 条产品线（设计 Owner），主导智慧楼宇实时报警系统重构；报警处理效率 **+70%**，满意度 **3.29 → 4.50**。",
    ]),
    ("海康威视", "交互设计工程师", "2017.06 — 2021.04", [
        "企业级客户端与智能硬件（IoT 产品 20+）的需求分析、信息架构与核心流程设计；iVMS-4200 相关模块页面 **-10%**。",
        "1600K 网络键盘解决 60 余项可用性问题，满意度 **78% → 89%**；获发明专利 1 项。",
    ]),
]

for co, role, date, bullets in JOBS:
    p = para(doc, before=4, after=0, line=1.15, keep_next=True)
    p.paragraph_format.tab_stops.add_tab_stop(CONTENT_W, WD_TAB_ALIGNMENT.RIGHT)
    c = p.add_run(co)
    set_font(c, 10.5, serif=True, color=INK)
    sep = p.add_run("   ")
    set_font(sep, 9)
    rl = p.add_run(role)
    set_font(rl, 8.6, color=MUTED)
    d = p.add_run("\t" + date)
    set_font(d, 7.8, color=MUTED, spacing=0.4)

    for b in bullets:
        bp = para(doc, before=1.2, after=0, line=1.27, left=0.42, hanging=0.42)
        dash = bp.add_run("–  ")
        set_font(dash, 8.6, color=MUTED)
        rich(bp, b, 8.7)

# ---------- 代表项目（STAR） ----------
section(doc, "03", "代表项目")

PROJECTS = [
    {
        "name": "华为 Answer 2.0 AI 办公｜企业级 Agent 工作伙伴",
        "meta": "产品经理 + 体验设计 · 双角色",
        "date": "",
        "rows": [
            ("S 背景", "信息散在 WeLink 文档、群聊、知识中心与代码仓库，跨系统办事要反复复制粘贴；知识源有 3 个入口，全量用户中仅约 700 人使用且极易选错。"),
            ("T 任务", "以 PD + UXD 双角色负责能力边界、产品骨架、状态化交互与反馈闭环；目标是让用户只记一个入口、敢把任务交出去、失败时知道为什么。"),
            ("A 行动", "确立「Chat 是入口、垂域是能力」，子代理封装为工具由主 Agent 调度；知识源三入口收敛为一个下拉，改由 Agent 消解冲突；状态写成枚举——无输出不出现复制/分享，失败不出现点赞；反馈改为两级标签并自动映射跟进人；内审推行「先画理想链路，再推架构」。"),
            ("R 结果", "覆盖 **近 1000 人**，试点团队满意度 **4.2**；入口 **3 → 1**；反馈可定位（7.1–7.20 共 85 条，BUG 22 条，按模块分派）；沉淀 Agent 设计体系 **45 个组件**。"),
        ],
    },
    {
        "name": "华为 Answer Agent｜企业级解决方案",
        "meta": "产品经理 · 0 → 1",
        "date": "2024.08 — 2024.11",
        "rows": [
            ("S 背景", "解决方案工程师大量重复编写方案，人工产出周期长；客户需求频繁变更，且有强合规、可溯源要求，人工模式稳定性差。"),
            ("T 任务", "作为 AI 产品负责人，负责需求定义、Agent 工作流、交互方案到灰度落地的全流程；目标是把产出周期压下来，同时守住合规底线。"),
            ("A 行动", "访谈真实用户梳理使用链路，区分适合 AI 生成与必须人工审核的环节；定义 Agent 任务拆解、记忆管理与版本回溯；与算法、合规、研发对齐输出边界并设置校验拦截；输出方案与规范，交付 Vibe Coding 的产品源码；设计灰度方案，按幻觉、准确性、使用门槛迭代交互。"),
            ("R 结果", "方案平均产出 **3 天 → 3 分钟**；试点 **12 位售前，10 位持续使用**；生成内容全部留痕可审计；项目获公司优胜奖（3 万元），个人获「价值之星」。"),
        ],
    },
    {
        "name": "锐捷 Source ID 3.0｜本体与数据产品",
        "meta": "产品设计师 · 模块体验 / Agent 链路",
        "date": "2026.04 — 2026.10",
        "rows": [
            ("S 背景", "产品包含本体、数据管道、数据管理等专业模块，用户需要理解数据与本体之间的关系并完成配置构建，专业门槛高、跨模块任务不连贯。"),
            ("T 任务", "负责三个核心模块的体验设计，并参与 Agent 任务链路与设计规范建设；重点是降低专业能力的使用门槛，改善跨模块任务的连贯性。"),
            ("A 行动", "梳理数据接入、加工、建模、管理的依赖关系，找出跨模块断点；用 Lovable 做原型探索，验证跨模块信息架构与关键操作路径；设计 Agent 任务链路（输入 → 任务理解 → 构建执行 → 结果确认），定义过程状态、异常反馈与人工介入；探索 Agent 辅助与手动操作的协同，避免复杂任务成为黑盒；沉淀任务流程与状态表达原则。"),
            ("R 结果", "形成面向企业级复杂任务的 Agent 体验设计实践，兼顾配置效率、过程透明与人工控制；跨模块路径统一，Agent 过程可见、可回退。"),
        ],
    },
    {
        "name": "华为 Talent 实训云平台",
        "meta": "体验设计师 · 产品线设计 Owner / 设计组长",
        "date": "2022.04 — 2026.04",
        "rows": [
            ("S 背景", "教务、教学、考评分散在多套系统，校方 / 教师 / 学生三类角色任务割裂，教师批阅与签到重复劳动多。"),
            ("T 任务", "负责平台的体验规划与设计落地，覆盖教务、教学、考评主链路，推动产品结构与关键流程迭代。"),
            ("A 行动", "结合用户研究、可用性测试与行为数据定位问题；用 KANO 与竞品分析排优先级，先做「教务—考评」闭环；按三类角色重排任务流，把高频操作前移到首屏；参与组件体系与设计规范建设。"),
            ("R 结果", "部署 **7 所职业院校**；批阅、签到等高频模块使用率 **+13%**；两次获「明日之星」。"),
        ],
    },
]

for i, pr in enumerate(PROJECTS, 1):
    p = para(doc, before=5.5, after=0, line=1.12, keep_next=True)
    p.paragraph_format.keep_together = True
    p.paragraph_format.tab_stops.add_tab_stop(CONTENT_W, WD_TAB_ALIGNMENT.RIGHT)
    idx = p.add_run(f"{i:02d}  ")
    set_font(idx, 7.2, color=MUTED, spacing=1.0)
    nm = p.add_run(pr["name"])
    set_font(nm, 10.4, serif=True, color=INK)
    sp = p.add_run("   ")
    set_font(sp, 9)
    mt = p.add_run(pr["meta"])
    set_font(mt, 7.8, color=MUTED, spacing=0.3)
    d = p.add_run("\t" + pr["date"])
    set_font(d, 7.8, color=MUTED, spacing=0.4)

    for ri, (label, text) in enumerate(pr["rows"]):
        rp = para(doc, before=1.0, after=0, line=1.27, left=1.15, hanging=1.15)
        rp.paragraph_format.keep_together = True
        rp.paragraph_format.tab_stops.add_tab_stop(Cm(1.15))
        lb = rp.add_run(label + "\t")
        set_font(lb, 7.4, color=MUTED, spacing=0.6)
        rich(rp, text, 8.7)

# ---------- 教育经历 ----------
section(doc, "04", "教育经历")
for school, degree in (("浙江工业大学", "设计学硕士"),
                       ("河南理工大学", "工业设计学士"),
                       ("河南理工大学", "会计学（双学位）")):
    p = para(doc, before=3, after=0, line=1.35)
    s = p.add_run(school)
    set_font(s, 9.4, serif=True, color=INK)
    sep = p.add_run("   ")
    set_font(sep, 9)
    g = p.add_run(degree)
    set_font(g, 8.6, color=MUTED)

p = para(doc, before=6, after=0, line=1.4)
rich(p, "设计学出身，本科还读了会计学双学位——既能把东西画清楚，也看得懂数字。", 8.2, color=MUTED)

# ---------- 专业技能 ----------
section(doc, "05", "专业技能")
SKILLS = [
    ("业务领域", "企业级 SaaS、ToB 中后台、IoT、本体建模、数据管道与数据治理、多角色权限与审批链路"),
    ("AI 与工具", "LLM 应用与 RAG、Agent 任务链路、意图澄清、结果校验；Lovable、Axure、Figma、Principle、C4D / AE"),
    ("方法论与输出", "Agent 设计规范 45 个组件、TECS-R 体验度量、六维设计走查；45+ 篇文章、公众号 3000+ 订阅"),
]
for label, text in SKILLS:
    p = para(doc, before=2.6, after=0, line=1.4, left=1.85, hanging=1.85)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(1.85))
    lb = p.add_run(label + "\t")
    set_font(lb, 7.4, color=MUTED, spacing=0.6)
    rich(p, text, 8.7)

p = para(doc, before=7, after=0, line=1.4)
rich(p, "文章被「人人都是产品经理」「优设网」「微交互」等平台推荐；六维设计走查体系在多条产品线落地。", 8.2, color=MUTED)

cp = doc.core_properties
cp.title = "黄红艺 · 产品经理 · 简历"
cp.author = "黄红艺"
cp.subject = "产品经理（AI 应用 / Agent / B 端平台）"

doc.save("黄红艺-简历-产品经理.docx")
print("saved")
