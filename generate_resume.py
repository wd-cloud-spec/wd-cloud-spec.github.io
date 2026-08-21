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


def project_block(doc, name, tagline, bullets, tags, star=False, scale_note=None):
    """项目经历块：项目名 + 一句话 + bullet + 标签行。"""
    # 项目名
    p = para(doc, space_before=8, space_after=2, line=1.2)
    title = ("★ " if star else "") + name
    r = p.add_run(title)
    set_font(r, size=10.5, bold=True, color=COLOR_ACCENT if star else COLOR_TEXT)
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
        p = para(doc, space_before=2, space_after=2, line=1.2)
        r = p.add_run("  |  ".join(tags))
        set_font(r, size=8.5, color=COLOR_TAG, italic=True)
    if scale_note:
        para(doc, scale_note, size=8.5, color=COLOR_TAG, space_after=6, line=1.2)


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
para(doc, "AI产品构建者  ·  如何用AI去提高效率", size=11,
     color=COLOR_ACCENT, space_after=4, line=1.2)

para(doc, "📧 1252395926@qq.com   📞 130 3283 9382   🔗 github.com/wd-cloud-spec   🔗 gitee.com/scyilang_0",
     size=9, color=COLOR_SUB, space_after=2, line=1.2)
para(doc, "📍 重庆 · 成都  ·  随时到岗  ·  30岁  ·  电子科技大学（985）本科",
     size=9, color=COLOR_SUB, space_after=8, line=1.2)

# ============ 一句话定位 ============
section_title(doc, "一句话定位")
para(doc, "7年跨界：平面设计3年（审美与用户体验）→ B端产品经理1年（企业痛点与方案输出）→ 研学创业3年（从0到1的完整商业闭环，近百场交付、零安全事故）→ 独立构建多个AI产品（把大模型能力变成可交付、可变现的应用）。",
     size=10, color=COLOR_TEXT, space_after=4, line=1.35)

# ============ 核心能力（三栏表格）============
section_title(doc, "核心能力")

table = doc.add_table(rows=1, cols=3)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = True

caps = [
    ("商业闭环", "创业3年，独立完成产品研发、渠道拓客、商务谈判、交付回款、客户维护全链路。会算账、能控本、懂风控。"),
    ("技术落地", "独立完成AI产品全栈开发：Python/Node.js/JavaScript/PySide6，RAG选型、OCR降级链、多Agent编排均有实战。"),
    ("设计能力", "平面设计3年，兼具B端产品与售前思维，从原型到视觉落地全流程，独立输出界面与投标物料。"),
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
    doc, "DeepSeek Harness Mobile — 手机远程指挥台（旗舰项目）",
    "DeepSeek Harness 上线48小时内完成插件开发，用DSH让DSH更方便使用：手机实时镜像 PC 端 AI Agent 工作流，远程发言、审批放行。",
    [
        "双面插件架构：宿主进程端 + 浏览器UI端同一包分发，~2,650行核心代码，一键安装脚本，MIT开源",
        "安全桥接：14路径白名单代理 + 旋转Token + 本机Origin校验，恶意页面无法跨站攻击",
        "原生 AbortSignal 三级捕获：从手机远程停止 PC 端 Agent 任务（穿透沙箱伪造实现）",
        "消息队列 + 90秒TTL：离线会话唤醒后自动送达；手机与PC对审批/问答首答优先",
        "零外部资源单文件UI：深色指挥台设计（墨蓝+琥珀+青），notched屏安全区适配",
    ],
    ["DeepSeek Harness", "SSE", "移动端", "安全桥接", "Agent控制"],
    star=True,
    scale_note="工程规模：~2,650行 · 双面插件 · 单文件移动端UI · 12条工程踩坑记录",
)

project_block(
    doc, "CodeClean — AI 代码清理平台（旗舰项目）",
    "一套把安全做到极致的 AI 代码清理平台：帮助不懂代码的人去检查冗余代码文件。清理端 7 Agent\"做手术\"，学习端 5 Agent\"不再犯同样的错\"。",
    [
        "三道防线：沙箱副本只读执行 + .quarantine/ 只隔离不删除（SHA256溯源一条命令回迁）+ 基线回归校验自动回滚",
        "Tree-sitter AST + 框架感知检测（FastAPI/Django/Vue 动态路由），从源头降低误删率",
        "L1/L2/L3 三级风险分流，L2 暂停人工审批（LangGraph checkpoint 断点恢复）",
        "十节点 LangGraph DAG 编排 + SqliteSaver 断点续跑，FATAL/RETRY/DEGRADE 三级异常韧性",
        "五阶段学习闭环：Harvester→Evaluator→Validator→Integrator→Monitor",
    ],
    ["Python", "LangGraph", "12 Agent", "Tree-sitter", "沙箱隔离"],
    star=True,
    scale_note="工程规模：5000行 Python · 36个模块 · 12个Prompt模板 · 85项自动化测试全过",
)

project_block(
    doc, "铜雀台 × 商单台 — AI 销售管理平台（旗舰项目）",
    "双子星架构：铜雀台管客户（CRM），商单台管履约（报价→合同→回款→交付→售后），同一 Supabase 账号体系数据天然互通；语音助手让\"说话\"变成录客户、开报价、建订单。",
    [
        "语音→动作流水线：MediaRecorder→16kHz PCM重采样→讯飞WS（手写HMAC-SHA256签名）→DeepSeek结构化意图→前端执行类型化动作",
        "50产品目录注入system prompt，口语化产品名自动映射SKU并计价（实测\"高配工位套餐2套\"自动计价21,897元）",
        "金额/状态一致性由 PostgreSQL RPC 保证；审批门控漏斗（won/lost回退需审批+stage_logs审计）",
        "版本化幂等迁移：从2人团队演化到20-50人组织（层级权限+公海自动回收）；Repository模式（Supabase云+IndexedDB离线缓存）",
    ],
    ["Next.js", "React", "Supabase", "DeepSeek", "讯飞ASR", "PWA"],
    star=True,
    scale_note="工程规模：两个应用 ~15,000行 · 已部署 Vercel · PWA可安装 · 真实可上线",
)

project_block(
    doc, "成绩统计工具 — AI×教育桌面应用",
    "现代风格设计完整工作台：AI视觉识别让教师从手工录入中解放，双击即用的桌面工具。",
    [
        "Qwen-VL 视觉大模型识别成绩表格→结构化JSON，支持多图合并按姓名归并",
        "4级OCR降级链路（Qwen-VL→腾讯云→RapidOCR→PaddleOCR）逐级兜底，有网用AI、无网照样干",
        "PySide6 + Fluent Design 界面 + 视频动态背景，教师零学习成本",
        "PyInstaller打包.exe，SQLite本地存储，数据不离开学校；4类图表分析+丢分关键词检索",
    ],
    ["Python", "Qwen-VL", "PySide6", "Fluent Design", "SQLite"],
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
para(doc, "个人主页：wd-cloud-spec.github.io（AI项目视频演示）  ·  GitHub：github.com/wd-cloud-spec  ·  随时到岗  ·  目标城市：重庆、成都",
     size=9.5, color=COLOR_SUB, space_after=2, line=1.3)

out = "王栋-简历.docx"
doc.save(out)
print(f"已生成：{out}")
