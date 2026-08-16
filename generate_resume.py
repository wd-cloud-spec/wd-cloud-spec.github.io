# -*- coding: utf-8 -*-
"""生成 王栋 的 Word 简历（AI产品构建者方向）。

用法：python generate_resume.py
输出：王栋-简历.docx
修改内容后重跑即可重新生成。
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------- 常量 ----------
FONT_CN = "微软雅黑"
FONT_EN = "Calibri"
COLOR_TEXT = RGBColor(0x33, 0x33, 0x33)
COLOR_SUB = RGBColor(0x6B, 0x65, 0x60)
COLOR_ACCENT = RGBColor(0x8B, 0x76, 0x40)   # 金棕
COLOR_LINE = RGBColor(0xC4, 0xA9, 0x75)     # 亮金
COLOR_TAG = RGBColor(0x5A, 0x50, 0x45)


def set_font(run, size=10, bold=False, color=COLOR_TEXT, italic=False, cn=FONT_CN):
    """设置中英文字体与字号。"""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = FONT_EN
    run._element.rPr.rFonts.set(qn("w:eastAsia"), cn)


def add_bottom_border(paragraph, color="C4A975", size="8"):
    """给段落加底部边框线。"""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def para(doc, text="", size=10, bold=False, color=COLOR_TEXT, italic=False,
         align=None, space_before=0, space_after=4, line=1.3, cn=FONT_CN):
    """添加段落并返回。"""
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if text:
        r = p.add_run(text)
        set_font(r, size=size, bold=bold, color=color, italic=italic, cn=cn)
    return p


def section_title(doc, text):
    """章节标题：加粗 + 金色下划线。"""
    p = para(doc, text, size=12, bold=True, color=COLOR_ACCENT,
             space_before=14, space_after=8, line=1.2)
    add_bottom_border(p)
    return p


def project_block(doc, name, tagline, bullets, tags, star=False):
    """项目经历块：项目名 + 一句话 + 3条bullet + 标签行。"""
    # 项目名 + 标签
    p = para(doc, space_before=8, space_after=2, line=1.2)
    title = ("★ " if star else "") + name
    r = p.add_run(title)
    set_font(r, size=10.5, bold=True, color=COLOR_TEXT)
    if star:
        r.font.color.rgb = COLOR_ACCENT
    # 一句话
    para(doc, tagline, size=9.5, color=COLOR_SUB, space_after=2, line=1.3)
    # bullets
    for b in bullets:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(1)
        pf.line_spacing = 1.3
        pf.left_indent = Cm(0.5)
        r = p.add_run("• " + b)
        set_font(r, size=9.5, color=COLOR_TEXT)
    # 标签行
    if tags:
        p = para(doc, space_before=2, space_after=6, line=1.2)
        r = p.add_run("  |  ".join(tags))
        set_font(r, size=8.5, color=COLOR_TAG, italic=True)


def job_block(doc, period, company_role, desc_lines):
    """工作经历块。"""
    p = para(doc, space_before=8, space_after=2, line=1.2)
    r = p.add_run(period + "  ")
    set_font(r, size=9.5, bold=True, color=COLOR_ACCENT)
    r2 = p.add_run(company_role)
    set_font(r2, size=10.5, bold=True, color=COLOR_TEXT)
    for d in desc_lines:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(1)
        pf.line_spacing = 1.3
        pf.left_indent = Cm(0.5)
        r = p.add_run("• " + d)
        set_font(r, size=9.5, color=COLOR_TEXT)


# ============================================================
doc = Document()

# 页面设置 A4
sec = doc.sections[0]
sec.page_width = Cm(21)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(1.8)
sec.bottom_margin = Cm(1.6)
sec.left_margin = Cm(1.8)
sec.right_margin = Cm(1.8)

# ============ 头部 ============
para(doc, "王栋", size=22, bold=True, space_after=2, line=1.1)
para(doc, "AI产品构建者  ·  能独立走通商业闭环的AI应用全栈搭建", size=11,
     color=COLOR_ACCENT, space_after=4, line=1.2)

para(doc, "📧 1252395926@qq.com   📞 130 3283 9382   🔗 gitee.com/scyilang_0   🌐 wd-cloud-spec.github.io",
     size=9, color=COLOR_SUB, space_after=2, line=1.2)
para(doc, "📍 重庆 · 成都  ·  随时到岗  ·  30岁  ·  电子科技大学（985）本科",
     size=9, color=COLOR_SUB, space_after=8, line=1.2)

# ============ 一句话定位 ============
section_title(doc, "一句话定位")
para(doc, "7年跨界：平面设计3年（审美与用户体验）→ B端产品经理1年（企业痛点与方案输出）→ 研学创业3年（从0到1的完整商业闭环，近百场交付、零安全事故）→ 独立构建4个AI产品（把大模型能力变成可交付、可变现的应用）。",
     size=10, color=COLOR_TEXT, space_after=4, line=1.35)

# ============ 核心能力（三栏表格）============
section_title(doc, "核心能力")

table = doc.add_table(rows=1, cols=3)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = True

caps = [
    ("商业闭环", "创业3年，独立完成产品研发、渠道拓客、商务谈判、交付回款、客户维护全链路。会算账、能控本、懂风控。"),
    ("技术落地", "独立完成AI产品全栈开发：Python/Node.js/PySide6/Electron，RAG选型、OCR降级链、多Agent编排均有实战。"),
    ("设计能力", "平面设计3年，项目执行中可独立完成界面设计与视觉物料，无需等待设计师，保证交付完整性。"),
]
for i, (title, desc) in enumerate(caps):
    cell = table.cell(0, i)
    cell.width = Cm(5.8)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.3
    r = p.add_run(title)
    set_font(r, size=10, bold=True, color=COLOR_ACCENT)
    p2 = cell.add_paragraph()
    p2.paragraph_format.line_spacing = 1.3
    r2 = p2.add_run(desc)
    set_font(r2, size=9, color=COLOR_TEXT)

# ============ AI项目经历 ============
section_title(doc, "AI 项目经历")

project_block(
    doc, "CodeClean — AI 代码清理平台（旗舰项目）",
    "一套把安全做到极致的代码清理平台：清理端7 Agent\"做手术\"，学习端5 Agent\"不再犯同样的错\"。",
    [
        "安全架构：沙箱副本只读执行 + .quarantine/ 只隔离不删除（SHA256溯源，一条命令回迁）+ 基线回归校验，异常自动回滚整个批次",
        "Tree-sitter AST 解析 + 框架感知检测（FastAPI/Django/Vue 动态路由），L1/L2/L3 三级风险分流，L2 暂停人工审批",
        "LangGraph 十节点 DAG 编排 + SqliteSaver 断点续跑，FATAL/RETRY/DEGRADE 三级异常韧性，零Docker、Windows原生",
        "工程规模：5000行 Python · 36个模块 · 12个Agent · 12个Prompt模板 · 85项自动化测试全过",
    ],
    ["Python", "LangGraph", "12 Agent", "Tree-sitter", "沙箱隔离", "学习闭环"],
    star=True,
)

project_block(
    doc, "企业管理系统 — 中小企业AI全栈管理平台",
    "Node.js 零依赖后端集成 OA/CRM/WMS 三大子系统，DeepSeek 能力嵌入 6 个真实业务流程，Electron 打包为桌面应用。",
    [
        "6大AI功能：智能填单（自然语言→表单）、发票文字识别、CRM跟进摘要、赢单预测、采购建议、全局AI助手",
        "自研中文NLP确定性后处理修正AI输出；纯手写XLSX解析器；100并发压测零崩溃、零5xx",
        "统一认证+RBAC三级权限（17项细粒度权限），Token会话，scrypt密码哈希，可局域网/私有云部署",
    ],
    ["Node.js", "DeepSeek", "Electron", "RBAC", "零依赖"],
)

project_block(
    doc, "成绩统计工具 — AI×教育垂直场景桌面应用",
    "面向中学教师的桌面端成绩管理：AI视觉识别让教师从手工录入中解放，有网无网都能用。",
    [
        "Qwen-VL 视觉大模型识别成绩表格→结构化JSON，支持多图合并按姓名归并；图像预处理流水线提升低质量照片识别率",
        "4级OCR降级链路（Qwen-VL→腾讯云→RapidOCR→PaddleOCR）逐级兜底，断网可用",
        "PySide6 + Fluent Design（Win11风格）界面，PyInstaller打包.exe，SQLite本地存储，双击即用",
    ],
    ["Python", "Qwen-VL", "PySide6", "OCR Pipeline", "SQLite"],
)

project_block(
    doc, "RAG知识库MAX — 私有化桌面端RAG系统",
    "面向招投标/法务/研究的本地知识库：17种模块化RAG策略、6家国产大模型、全数据AES-256加密。",
    [
        "17种RAG策略三层架构（切分组6选1+增强组6选多+高阶组5选多），实时冲突校验，智能/专家双模式",
        "bge-m3本地Embedding+bge-reranker精排，向量永久不变、仅增量重建；离线模式自动降级（CRAG→Self-RAG等）",
        "硬件自适应量化（FP16/8bit/4bit），多密钥轮询+故障自动切换，Token预算管理",
    ],
    ["Python", "bge-m3", "GraphRAG", "PyQt6", "AES-256"],
)

# ============ 工作经历 ============
section_title(doc, "工作经历")

job_block(
    doc, "2023.08 - 至今",
    "高校研学项目 · 创始人 / 项目负责人（自主创业）",
    [
        "独立完成产品研发→渠道拓客→商务谈判→项目交付→客户维护→复购的完整商业闭环，深耕成都高校研学市场",
        "累计落地近百场研学实践，师生满意度95%以上，零安全事故、零投诉",
        "独立核算成本利润、统筹团队排班培训，多场次项目同步有序落地",
    ],
)

job_block(
    doc, "2022.05 - 2023.08",
    "信锐科技 · B端产品经理（政企/高校园区网络）",
    [
        "对接政企与高校客户，挖掘网络建设与信息化改造需求，撰写解决方案、技术白皮书与招投标资料",
        "跟进产品迭代与项目落地，联动技术、实施团队解决执行中的对接问题",
    ],
)

job_block(
    doc, "2019.01 - 2022.05",
    "淘宝美工 / 平面设计师",
    [
        "3年全品类视觉设计：电商店铺装修、品牌画册、活动物料、招投标画册等全案输出",
        "具备独立完成产品界面设计的能力，保证项目视觉交付不依赖外部设计师",
    ],
)

# ============ 教育背景 ============
section_title(doc, "教育背景")
para(doc, "电子科技大学（985/211） · 电波传播与天线专业 · 本科  |  2014.09 - 2018.06",
     size=10, color=COLOR_TEXT, space_after=2, line=1.3)

# ============ 附加信息 ============
section_title(doc, "附加信息")
para(doc, "个人主页：wd-cloud-spec.github.io（AI项目视频演示）  ·  随时到岗  ·  目标城市：重庆、成都",
     size=9.5, color=COLOR_SUB, space_after=2, line=1.3)

out = "王栋-简历.docx"
doc.save(out)
print(f"已生成：{out}")
