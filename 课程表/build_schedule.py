#!/usr/bin/env python3
"""把三天实训提纲整理成 Excel 课程表。

原文八列逐字录入。天数、模块、任务形式是便于排课的归纳字段：
第 1 天 L1–L8（以 L8 的 Day1 作品收口），第 2 天 L9–L16（以 L16 登记 Day2 作品），
第 3 天 L17–L24。提纲没有给出每节分钟数和钟点，表里不排时间。
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.pagebreak import Break

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "三天实训课程表.xlsx"

FONT = "微软雅黑"
INK = "1C2B33"
MUTED = "5E6A72"
CREAM = "F7F4EF"
WHITE = "FFFFFF"
LINE = "E2E8EC"
GOLD = "C4A574"
INPUT = "FBF6EE"
GREEN = "DCEFE4"
GREEN_INK = "1E6B45"

# 课时、主题、教学目标、教学内容要点、教学方法、学员任务、产出、考核标准
# 与提纲原文一致。
SOURCE = """
L1	开营与工具总览	能独立完成 Codex 登录与 dsh 本地启动，并说清 Cursor、WorkBuddy、Claude Code 与 Codex、dsh 各自的定位。	三天地图与验收规则；Cursor、WorkBuddy、Claude Code 各自擅长、优劣势与国内可用性，三者与 Codex、dsh 的关系	讲授 · 投屏演示 · 学员操作 · 检核	学员完成 Codex 登录、dsh 启动，填环境检查表与工具速查卡	环境检查表 + 工具速查卡	Codex 能发指令，dsh 本地页面能打开
L2	Codex 主线	能用 Codex 完成第一个文件任务，并解释它读了哪些文件、跑了什么命令、改了什么。	CLI、Agent、Skill、MCP 四个基础概念；工作区与任务；多 API key 配置位置与安全习惯	讲授 · 演示 · 实操 · 检核	学员完成自己的第一个文件任务	任务记录 + 一句验收说明	能解释「我让 Codex 做了什么、结果怎么检查」
L3	DeepSeek Harness 主线	能在 dsh 里复现同一任务，并判断哪些任务更适合交给 dsh、哪些交给 Codex。	dsh 的一切皆插件、Skills 机制、与 Codex 的分工；开发者预览阶段的注意事项；多 API key 接入	讲授 · 演示 · 实操 · 对比讨论	学员在 dsh 里复现该任务	dsh 任务记录 + 三句话对比笔记	能说出「什么任务更愿意交给 Codex、什么更愿意交给 dsh」
L4	生图工具矩阵	能按「要艺术感 / 要文字准 / 要本地免费」选出合适的生图模型，并做出 4 图对比板。	Midjourney、即梦、GPT Image 2.5、Nano 2.1、z-image、qwen-image、开源模型各自风格、文字排版、编辑能力	讲授 · 多模型演示 · 实操 · 检核	学员用至少两个模型出同一主题图	4 图对比板（标注模型、提示词、耗时）	能说出「想要艺术感找谁、想要文字准找谁、想本地免费找谁」
L5	生视频工具矩阵	能按首帧、时长、预算三个判断点选视频模型，并产出双模型对比片段。	Seedance、MiniMax H3、Wan 3.0、gork 等的首帧控制、时长、一致性、价格档	讲授 · 双模型演示 · 实操 · 检核	学员用同一首帧跑一个模型	双模型对比片段或一段成品	能说出选视频模型的三个判断点（首帧、时长、预算）
L6	音频与音乐	能在在线编曲/配乐与本地配音之间按场景取舍，并交付一段配乐与一句中文配音。	编曲 Suno、oiioii；配乐 MiniMax；本地配音 indexTTS 2.5，各自适用场景与音质取舍	讲授 · 现场生成演示 · 实操 · 检核	学员生成自己的配乐与配音	配乐文件 + 配音文件	能说清「什么时候用在线、什么时候用本地 indexTTS」
L7	国内平台接入与成本	能算出 15 秒广告的模型成本构成，并填出一张平台接入与成本对照表。	libtv、tapnow、shotlabs、即梦、小云雀、海螺AI 怎么接入、适合做什么、充值方式；四类计费（订阅 / 积分 / 按秒 / 开源部署）	讲授 · 平台实操 · 现场填表 · 检核	学员现场查最新价格，填成本对比表（最新价以官网为准）	平台接入与成本对照表	能算出一个 15 秒广告的模型成本大致构成
L8	实操整合	能把图、视频、音频串成一个任务，完成「海报 + 视频 + 配音」小闭环。	多 API key 的安全管理；一个任务串起图、视频、音频的调用顺序	讲授 · 流程演示 · 项目制实操 · 检核	学员完成自己的小闭环	Day1 作品（1 图 + 1 视频 + 1 配音）	三件产出齐备，路径可复现
L9	爆款短剧/短视频拆解	能逐段拆解 3 条爆款内容，说清前 3 秒为什么留人，并提炼可迁移规律。	钩子、节奏、情绪、冲突、行动引导；抖音、视频号、小红书差异	讲授 · 现场拆解演示 · 实操 · 检核	学员拆 3 条内容	爆款拆解卡 3 份 + 可迁移规律	能说出「这条为什么在前 3 秒留住人」
L10	剧本与脚本创作	能写出有钩子、有冲突、有行动引导的 30–60 秒脚本，并拆出分镜清单。	选题、一句话梗概、单集结构、系列钩子；对照六百斤剧本讲结构	讲授 · 范例演示 · 实操 · 检核	学员写自己的脚本并拆分镜	脚本 + 分镜清单	脚本有钩子、有冲突、有行动引导，可直接进生产
L11	美术资产设计与优化	能产出 3 张可跨镜头复用的资产卡，并说明资产编号与一致性锁的用法。	角色卡、场景卡、道具卡、一致性锁、资产编号；对照六百斤资产冻结库	讲授 · 资产库演示 · 实操 · 检核	学员出自己脚本的资产卡	资产卡 3 张 + 编号说明	角色与场景可跨镜头复用不跑形
L12	视频提示词结构与优化	能写出结构完整、时间码清晰的视频提示词 3 条，并能按问题逐项微调。	主体、动作、景别、运镜、光线、负面词、首帧锁定；时间码提示词写法	讲授 · 同句三改演示 · 实操 · 检核	学员写并优化自己的 3 条提示词	视频提示词 3 条（含首帧图）	提示词可直接投喂，且知道每段怎么微调
L13	基础剪辑逻辑（上）	能完成节奏成立、废段已去、故事顺序清楚的 30 秒粗剪。	粗剪、节奏、气口、卡点；素材分类与时间线整理	讲授 · 实剪演示 · 实操 · 检核	学员完成自己的粗剪	粗剪样片	节奏不拖，废段已去，故事顺序成立
L14	基础剪辑逻辑（下）	能完成音画同步、字幕无错、规格合规的成片，并留下质检记录。	字幕、音效、转场、画幅、封面、导出规格与质检清单	讲授 · 精剪演示 · 实操 · 质检互查	学员完成精剪导出	成片 1 条 + 质检记录	音画同步、字幕无错、规格符合平台
L15	AI 广告商业实战案例	能做出同一主体三画幅不跑形的橱窗物料，并说明每个画幅的用途。	B36、小米音响、sony-optics 三个广告案例的全链路复盘（简报、分镜、素材、成片、交付）	讲授 · 案例复盘 · 项目制实操 · 检核	学员完成自己的 3 画幅物料或 15 秒分镜加首帧	广告物料	同一主体三画幅不跑形，能说清每画幅用途
L16	当日复盘	能按「能看懂、能复用、有亮点」三句话互评，并登记作品清单与改进项。	作品互评三句话标准（能看懂、能复用、有亮点）	互评研讨 · 教师点评 · 登记	同桌互评，登记 Day2 作品	作品清单 + 3 条改进	每件作品都被看过、都有改进项
L17	AI 短剧案例（上）：剧本到资产冻结	能产出编号一致、可被后续镜头引用的剧本与资产注册表。	六百斤案例的文学细化、资产注册表、编号锁定	讲授 · 案例演示 · 分组实操 · 检核	每组定剧本、做资产注册表	剧本 + 资产表	角色、场景、道具编号一致，可被后续镜头引用
L18	AI 短剧案例（中）：分镜与提示词	能完成一场戏的分镜、首帧与可直接投喂的提示词，镜头衔接不跳戏。	分镜结构、首帧锁定、Seedance 提示词、多人站位（multi-casting-protocol）	讲授 · 分镜演示 · 分组实操 · 检核	每组完成自己那场戏	分镜 + 首帧 + 生成提示词	提示词可直接投喂，镜头衔接不跳戏
L19	AI 短剧案例（下）：成片与交付	能拼出一段 30 秒样片，并打包成可复现的交付包。	片段生成、合成、成片质检、交付包四件套	讲授 · 合成演示 · 分组实操 · 质检	每组完成样片与交付包	短剧样片 + 交付包	音画可看、复现步骤可照做、成本有记录
L20	AI 网站到全案策划案例	能做出一页可打开的落地页草案，并写出含策略、内容、视觉、投放的全案一页纸。	yiwu 站、iphone-duo 视觉包、knowledge-site 拆解；网站如何配合全案策划	讲授 · 站点拆解演示 · 实操 · 检核	学员做自己的页面草案	页面 + 全案一页纸	页面能打开，全案纸写清策略、内容、视觉、投放四块
L21	项目适配与商业报价	能开出含工作量、模型成本、改版条款、交付物清单的报价单。	按项目选工具、算成本、报价格；商业报价表与 opus-video-router 的档位思路	讲授 · 报价演示 · 实操 · 互评	学员给自己的作品报价	报价单	报价单含工作量、模型成本、改版条款、交付物清单
L22	自建数据库与 Skill	能把一条顺手流程封装成 Skill，换机按说明安装即可跑通。	提示词库、资产库、案例库怎么搭；skill-creator 建包、SKILL.md 写法、skill-vetter 自审	讲授 · 封装演示 · 实操 · 换机验证	学员封装自己的 Skill	自建 Skill 包 + 数据库目录	换机能按说明安装跑通
L23	商业 SOP 工作流	能画出一张每步都有工具、Skill 与验收点的商业 SOP 图。	一站式 SOP：接单、简报、生产、质检、交付、复盘；定制 Skill 应用领域与课程命名（海报、电商、短剧、网站、MV、口播、企业 SOP）	讲授 · 拆解演示 · 实操 · 互评	学员画自己的 SOP 并标注每步用哪个工具与 Skill	SOP 一页纸	SOP 里每步都有工具、Skill 与验收点
L24	结业路演与证书	能按「输入、过程、结果、验收、成本」五段完成 5 分钟路演并通过评分。	5 分钟路演结构（输入、过程、结果、验收、成本）	路演答辩 · 评委评分 · 颁证	每组路演，评委按评分卡打分	评分结果 + FED 证书（价值 1,000 元，已含学费）	讲清输入、过程与验收，交付包当场可抽查打开
""".strip()

MODULES = {
    "L1": "工具主线",
    "L2": "工具主线",
    "L3": "工具主线",
    "L4": "生成矩阵",
    "L5": "生成矩阵",
    "L6": "生成矩阵",
    "L7": "成本闭环",
    "L8": "成本闭环",
    "L9": "拆解与脚本",
    "L10": "拆解与脚本",
    "L11": "资产与提示词",
    "L12": "资产与提示词",
    "L13": "剪辑成片",
    "L14": "剪辑成片",
    "L15": "广告实战",
    "L16": "当日复盘",
    "L17": "短剧全链路",
    "L18": "短剧全链路",
    "L19": "短剧全链路",
    "L20": "网站与报价",
    "L21": "网站与报价",
    "L22": "能力沉淀",
    "L23": "能力沉淀",
    "L24": "结业路演",
}

DAYS = {
    1: {
        "name": "第 1 天",
        "theme": "工具与生成",
        "path": "开营对齐 → Codex 与 dsh → 生图、生视频、音频 → 平台与成本 → 海报、视频、配音小闭环",
        "close": "Day1 作品（1 图 + 1 视频 + 1 配音）",
        "accent": "1F4E5F",
        "bg": "E7F2F5",
        "alt": "F4FAFB",
    },
    2: {
        "name": "第 2 天",
        "theme": "内容与成片",
        "path": "爆款拆解 → 脚本与分镜 → 资产卡 → 视频提示词 → 粗剪与精剪 → 广告三画幅 → 互评登记",
        "close": "作品清单 + 3 条改进",
        "accent": "8C5A3C",
        "bg": "F8F1EA",
        "alt": "FDF9F5",
    },
    3: {
        "name": "第 3 天",
        "theme": "案例与结业",
        "path": "短剧从剧本到样片交付 → 落地页与全案 → 商业报价 → 自建 Skill → 商业 SOP → 五分钟路演",
        "close": "评分结果 + FED 证书（价值 1,000 元，已含学费）",
        "accent": "2E5A45",
        "bg": "E8F3EC",
        "alt": "F5FBF7",
    },
}

NOTE = (
    "整理说明：①「天数」「模块」「任务形式」是为了排课加上的归纳，不属于原文八列。"
    "第 1 天为 L1–L8（原文以 L8 的 Day1 作品收口），第 2 天为 L9–L16（原文以 L16 登记 Day2 作品），"
    "第 3 天为 L17–L24（短剧案例至结业路演）。② 教学目标、教学内容要点、教学方法、学员任务、产出、考核标准均按提纲原文录入，未改写。"
    "③ 提纲未给出每节分钟数和上课钟点，因此表中不排时间。④ 工具与模型名称按原文保留（含 gork、oiioii、libtv 等写法）。"
)


def load_lessons() -> list[dict]:
    lessons = []
    for line in SOURCE.splitlines():
        parts = line.split("\t")
        if len(parts) != 8:
            raise SystemExit(f"原文列数不是 8：{parts[0] if parts else line}")
        lid, topic, objective, content, method, task, output, standard = parts
        index = int(lid[1:])
        if not 1 <= index <= 24:
            raise SystemExit(lid)
        day = 1 if index <= 8 else 2 if index <= 16 else 3
        lessons.append(
            {
                "id": lid,
                "day": day,
                "module": MODULES[lid],
                "topic": topic,
                "objective": objective,
                "content": content,
                "method": method,
                "task": task,
                "output": output,
                "standard": standard,
                "form": task_form(task),
            }
        )
    if [item["id"] for item in lessons] != [f"L{i}" for i in range(1, 25)]:
        raise SystemExit("课时编号不连续")
    return lessons


def task_form(task: str) -> str:
    if task.startswith("每组"):
        return "分组"
    if task.startswith("同桌"):
        return "同桌"
    if task.startswith("学员"):
        return "个人"
    raise SystemExit(f"无法判断任务形式：{task}")


def font(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def side(style="thin", color=LINE) -> Side:
    return Side(style=style, color=color)


def border(color=LINE, top=None, bottom=None) -> Border:
    return Border(
        left=side(color=color),
        right=side(color=color),
        top=top or side(color=color),
        bottom=bottom or side(color=color),
    )


LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def text_lines(text: str, width: float) -> int:
    usable = max(width - 2.4, 6)
    lines = 1
    used = 0.0
    for ch in str(text or ""):
        if ch == "\n":
            lines += 1
            used = 0.0
            continue
        weight = 2.05 if ord(ch) > 127 else 1.0
        if used + weight > usable and used > 0:
            lines += 1
            used = weight
        else:
            used += weight
    return lines


def apply_page(ws, title_rows: str, fit_height: int = 0, zoom: int = 110):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.45
    ws.page_margins.right = 0.45
    ws.page_margins.top = 0.72
    ws.page_margins.bottom = 0.55
    ws.page_margins.header = 0.28
    ws.page_margins.footer = 0.28
    ws.print_title_rows = title_rows
    ws.oddHeader.left.text = "三天实训课程表"
    ws.oddHeader.left.font = FONT
    ws.oddHeader.left.size = 10
    ws.oddHeader.right.text = ws.title
    ws.oddHeader.right.font = FONT
    ws.oddHeader.right.size = 10
    ws.oddFooter.left.text = "天数与模块为整理归纳；正文八列与提纲一致。未编钟点。"
    ws.oddFooter.left.font = FONT
    ws.oddFooter.left.size = 9
    ws.oddFooter.right.text = "第 &P 页 / 共 &N 页"
    ws.oddFooter.right.font = FONT
    ws.oddFooter.right.size = 9
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = zoom
    ws.sheet_properties.tabColor = "1F4E5F"
    ws.page_setup.horizontalDpi = 300
    ws.page_setup.verticalDpi = 300
    ws.sheet_format.defaultRowHeight = 20


def write_merged(ws, row, start, end, value, cell_font, cell_fill, align, height=None):
    for col in range(start, end + 1):
        cell = ws.cell(row, col, value if col == start else None)
        cell.font = cell_font
        cell.fill = cell_fill
        cell.alignment = align
    if end > start:
        ws.merge_cells(start_row=row, start_column=start, end_row=row, end_column=end)
    if height:
        ws.row_dimensions[row].height = height
    return ws.cell(row, start)


def link_to(cell, location: str, tooltip: str, cell_font):
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=location, tooltip=tooltip)
    cell.font = Font(
        name=cell_font.name or FONT,
        size=cell_font.size,
        bold=cell_font.bold,
        italic=cell_font.italic,
        color=cell_font.color,
        underline="single",
    )


def style_header_row(ws, row, headers, fills=None):
    ws.row_dimensions[row].height = 24
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row, col, header)
        cell.font = font(10, bold=True, color=WHITE)
        cell.fill = fill((fills or {}).get(col, "1F4E5F"))
        cell.alignment = CENTER
        cell.border = Border(
            left=side(color="1F4E5F"),
            right=side(color="1F4E5F"),
            top=side(color="1F4E5F"),
            bottom=side(style="medium", color=GOLD),
        )


def set_widths(ws, widths):
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width


def build_full(wb, lessons):
    ws = wb.create_sheet("完整课程表")
    widths = [14, 16, 10, 24, 40, 46, 26, 34, 32, 36, 12]
    headers = [
        "天数",
        "模块",
        "课时",
        "主题",
        "教学目标",
        "教学内容要点",
        "教学方法",
        "学员任务",
        "产出",
        "考核标准",
        "任务形式",
    ]
    last_col = len(headers)
    write_merged(
        ws, 1, 1, last_col,
        "三天实训课程表",
        font(20, bold=True, color=WHITE),
        fill(INK),
        Alignment(horizontal="left", vertical="center", indent=1),
        34,
    )
    write_merged(
        ws, 2, 1, last_col,
        "24 课时。天数、模块、任务形式为整理字段：第 1 天 L1–L8，第 2 天 L9–L16，第 3 天 L17–L24。"
        "底色对应三天：蓝、褐、绿。教学目标至考核标准与提纲原文一致。可用表头筛选某一天、某一模块或分组课。提纲未给出钟点，表中不排时间。",
        font(9, color=INK),
        fill(CREAM),
        Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1),
        32,
    )
    style_header_row(ws, 3, headers)
    ws.cell(3, 1).comment = Comment(
        "按提纲中的 Day1 作品（L8）和 Day2 作品（L16）整理，每天 8 节。提纲没有写出上课钟点。",
        "课程表",
        width=260,
        height=70,
    )
    ws.cell(3, 2).comment = Comment(
        "模块按授课区块归纳，便于筛选，不是提纲原列。",
        "课程表",
        width=220,
        height=50,
    )
    ws.cell(3, 11).comment = Comment(
        "由学员任务首词归纳：学员→个人，同桌→同桌，每组→分组。",
        "课程表",
        width=240,
        height=55,
    )

    center_cols = {1, 2, 3, 11}
    locations = {}
    for offset, lesson in enumerate(lessons):
        row = 4 + offset
        locations[lesson["id"]] = row
        day = DAYS[lesson["day"]]
        bg = day["bg"] if offset % 2 == 0 else day["alt"]
        values = [
            day["name"],
            lesson["module"],
            lesson["id"],
            lesson["topic"],
            lesson["objective"],
            lesson["content"],
            lesson["method"],
            lesson["task"],
            lesson["output"],
            lesson["standard"],
            lesson["form"],
        ]
        lines = 1
        top = side("medium", day["accent"]) if lesson["id"] in {"L1", "L9", "L17"} else None
        for value, width, col in zip(values, widths, range(1, last_col + 1)):
            cell = ws.cell(row, col, value)
            bold = col in (2, 3, 4, 9)
            color = day["accent"] if col in (1, 2, 3, 11) else INK
            cell.font = font(10, bold=bold, color=color)
            cell.fill = fill(bg)
            cell.alignment = CENTER if col in center_cols else LEFT
            cell.border = border(top=top)
            lines = max(lines, text_lines(str(value), width))
        ws.row_dimensions[row].height = max(48, lines * 18 + 12)

    ws.row_dimensions[28].height = 12
    note_row = 29
    write_merged(
        ws, note_row, 1, last_col, NOTE,
        font(8, color=MUTED),
        fill(WHITE),
        Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1),
        48,
    )
    set_widths(ws, widths)
    ws.freeze_panes = "E4"
    ws.auto_filter.ref = "A3:K27"
    apply_page(ws, "1:3", zoom=100)
    ws.sheet_properties.tabColor = "1F4E5F"
    ws.print_area = f"A1:K{note_row}"
    ws.oddHeader.right.text = "完整课程表"
    note_lines = text_lines(NOTE, sum(widths))
    ws.row_dimensions[note_row].height = max(48, note_lines * 14 + 8)
    return ws, locations


def build_source(wb, lessons):
    ws = wb.create_sheet("原文八列")
    headers = ["课时", "主题", "教学目标", "教学内容要点", "教学方法", "学员任务", "产出", "考核标准"]
    widths = [10, 28, 42, 48, 28, 36, 34, 38]
    write_merged(
        ws, 1, 1, 8, "原文八列",
        font(20, bold=True, color=WHITE), fill(INK),
        Alignment(horizontal="left", vertical="center", indent=1), 34,
    )
    write_merged(
        ws, 2, 1, 8,
        "以下八列与授课提纲一致，未增删字句，便于和「完整课程表」对照。天数、模块、任务形式不在本表。底色同为三天：蓝、褐、绿。",
        font(9, color=INK), fill(CREAM),
        Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1), 28,
    )
    style_header_row(ws, 3, headers)
    for offset, lesson in enumerate(lessons):
        row = 4 + offset
        day = DAYS[lesson["day"]]
        bg = day["bg"] if offset % 2 == 0 else day["alt"]
        values = [
            lesson["id"], lesson["topic"], lesson["objective"], lesson["content"],
            lesson["method"], lesson["task"], lesson["output"], lesson["standard"],
        ]
        lines = 1
        top = side("medium", day["accent"]) if lesson["id"] in {"L1", "L9", "L17"} else None
        for col, value in enumerate(values, 1):
            cell = ws.cell(row, col, value)
            cell.font = font(10, bold=col in (1, 2), color=day["accent"] if col == 1 else INK)
            cell.fill = fill(bg)
            cell.alignment = CENTER if col == 1 else LEFT
            cell.border = border(top=top)
            lines = max(lines, text_lines(value, widths[col - 1]))
        ws.row_dimensions[row].height = max(48, lines * 18 + 12)
    set_widths(ws, widths)
    ws.auto_filter.ref = "A3:H27"
    ws.freeze_panes = "C4"
    ws.print_area = "A1:H27"
    apply_page(ws, "1:3", zoom=100)
    ws.sheet_properties.tabColor = "8C5A3C"
    ws.oddHeader.right.text = "原文八列"
    return ws


def build_check(wb, lessons):
    ws = wb.create_sheet("产出与考核")
    headers = ["天数", "课时", "主题", "任务形式", "产出", "考核标准", "完成情况", "备注"]
    widths = [14, 10, 28, 12, 36, 42, 14, 28]
    write_merged(
        ws, 1, 1, 8, "产出与考核",
        font(20, bold=True, color=WHITE), fill(INK),
        Alignment(horizontal="left", vertical="center", indent=1), 34,
    )
    write_merged(
        ws, 2, 1, 6,
        "按节勾选学员交付。完成情况下拉选择「未开始」或「已完成」，已完成显示为绿色。备注留给现场。",
        font(9, color=INK), fill(CREAM),
        Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1), 28,
    )
    progress_label = ws.cell(2, 7, "完成进度")
    progress_label.font = font(9, bold=True, color=WHITE)
    progress_label.fill = fill("2E5A45")
    progress_label.alignment = CENTER
    progress = ws.cell(2, 8, '=COUNTIF(G4:G27,"已完成")&" / 24 节"')
    progress.font = font(12, bold=True, color="2E5A45")
    progress.fill = fill("E8F3EC")
    progress.alignment = CENTER
    header_fills = {7: "8C5A3C", 8: "8C5A3C"}
    style_header_row(ws, 3, headers, header_fills)
    ws.cell(3, 7).comment = Comment(
        "点开下拉，选择未开始或已完成。也可以留空。",
        "课程表",
        width=220,
        height=48,
    )
    ws.cell(3, 8).comment = Comment("现场记录：缺交、需返工、优秀样例等。", "课程表", width=220, height=48)

    for offset, lesson in enumerate(lessons):
        row = 4 + offset
        day = DAYS[lesson["day"]]
        bg = day["bg"] if offset % 2 == 0 else day["alt"]
        values = [
            day["name"], lesson["id"], lesson["topic"], lesson["form"],
            lesson["output"], lesson["standard"], None, None,
        ]
        lines = 1
        top = side("medium", day["accent"]) if lesson["id"] in {"L1", "L9", "L17"} else None
        for col, value in enumerate(values, 1):
            cell = ws.cell(row, col, value)
            if col in (7, 8):
                cell.fill = fill(INPUT)
                cell.font = font(10, color=INK)
                cell.alignment = CENTER if col == 7 else LEFT
            else:
                bold = col in (2, 3, 5)
                color = day["accent"] if col in (1, 2, 4) else INK
                cell.font = font(10, bold=bold, color=color)
                cell.fill = fill(bg)
                cell.alignment = CENTER if col in (1, 2, 4) else LEFT
            cell.border = border(top=top)
            lines = max(lines, text_lines(value or "", widths[col - 1]))
        ws.row_dimensions[row].height = max(42, lines * 18 + 10)

    dv = DataValidation(
        type="list",
        formula1='"未开始,已完成"',
        allow_blank=True,
        showErrorMessage=True,
        showInputMessage=True,
        errorTitle="完成情况",
        error="请从列表选择「未开始」或「已完成」，或留空。",
        promptTitle="完成情况",
        prompt="选择未开始或已完成。已完成的格子会变成绿色。",
    )
    dv.add("G4:G27")
    ws.add_data_validation(dv)
    ws.conditional_formatting.add(
        "G4:G27",
        CellIsRule(
            operator="equal",
            formula=['"已完成"'],
            fill=fill(GREEN),
            font=font(10, bold=True, color=GREEN_INK),
        ),
    )
    ws.conditional_formatting.add(
        "G4:G27",
        CellIsRule(
            operator="equal",
            formula=['"未开始"'],
            fill=fill("F4F1EA"),
            font=font(10, color=MUTED),
        ),
    )
    set_widths(ws, widths)
    ws.auto_filter.ref = "A3:H27"
    ws.freeze_panes = "D4"
    ws.print_area = "A1:H27"
    apply_page(ws, "1:3")
    ws.sheet_properties.tabColor = "2E5A45"
    ws.oddHeader.right.text = "产出与考核"
    return ws


def build_map(wb, lessons, locations):
    ws = wb.create_sheet("三日地图", 0)
    widths = [12, 28, 16, 12, 38, 42]
    headers = ["课时", "主题", "模块", "形式", "本节产出", "过关标准"]
    set_widths(ws, widths)
    write_merged(
        ws, 1, 1, 6, "三天实训课程表",
        font(22, bold=True, color=WHITE), fill(INK),
        Alignment(horizontal="left", vertical="center", indent=1), 38,
    )
    write_merged(
        ws, 2, 1, 6,
        "24 课时 · 每天 8 节 · 先看当天路径，再进完整课程表看授课明细。分组课在第 3 天的短剧三节和结业路演，第 2 天复盘为同桌互评。",
        font(10, color=INK), fill(CREAM),
        Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1), 30,
    )

    jumps = [
        (1, "完整课程表", "'完整课程表'!A1", "打开可筛选的 24 节明细"),
        (3, "产出与考核", "'产出与考核'!A1", "打开勾选表"),
        (5, "原文八列", "'原文八列'!A1", "打开未经归纳的提纲原文"),
    ]
    ws.row_dimensions[3].height = 22
    for col in range(1, 7):
        cell = ws.cell(3, col, None)
        cell.fill = fill(WHITE)
    for col, title, location, tip in jumps:
        cell = ws.cell(3, col, "查看" + title)
        link_to(cell, location, tip, font(10, bold=True, color="1F4E5F"))
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell.fill = fill("E7F2F5")
        ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 1)

    row = 5
    blocks = {1: lessons[0:8], 2: lessons[8:16], 3: lessons[16:24]}
    for day_no, items in blocks.items():
        day = DAYS[day_no]
        banner = (
            f"{day['name']}  ·  {day['theme']}\n"
            f"路径：{day['path']}\n"
            f"当日收口：{day['close']}"
        )
        for col in range(1, 7):
            cell = ws.cell(row, col, banner if col == 1 else None)
            cell.font = font(11, bold=True, color=WHITE)
            cell.fill = fill(day["accent"])
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        banner_lines = text_lines(banner, sum(widths))
        ws.row_dimensions[row].height = max(76, banner_lines * 18 + 16)
        row += 1
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row, col, header)
            cell.font = font(9, bold=True, color=day["accent"])
            cell.fill = fill(day["bg"])
            cell.alignment = CENTER
            cell.border = Border(bottom=side(style="thin", color=day["accent"]))
        ws.row_dimensions[row].height = 20
        row += 1
        first_data = row
        for offset, lesson in enumerate(items):
            bg = day["bg"] if offset % 2 == 0 else day["alt"]
            values = [
                lesson["id"], lesson["topic"], lesson["module"], lesson["form"],
                lesson["output"], lesson["standard"],
            ]
            lines = 1
            for col, value in enumerate(values, 1):
                cell = ws.cell(row, col, value)
                color = day["accent"] if col in (1, 3, 4) else INK
                if col == 4 and lesson["form"] == "分组":
                    color = "8C5A3C"
                cell.font = font(10, bold=col in (1, 2), color=color)
                cell.fill = fill(bg)
                cell.alignment = CENTER if col in (1, 3, 4) else LEFT
                cell.border = border("E6EEF1")
                lines = max(lines, text_lines(value, widths[col - 1]))
            link_to(
                ws.cell(row, 1),
                f"'完整课程表'!C{locations[lesson['id']]}",
                f"在完整课程表中查看 {lesson['id']}",
                font(10, bold=True, color=day["accent"]),
            )
            ws.cell(row, 1).alignment = CENTER
            ws.row_dimensions[row].height = max(34, lines * 18 + 10)
            row += 1
        box(ws, first_data - 2, 1, row - 1, 6, day["accent"])
        if day_no < 3:
            ws.row_dimensions[row].height = 16
            for col in range(1, 7):
                ws.cell(row, col).fill = fill(WHITE)
            ws.row_breaks.append(Break(id=row))
        row += 1

    counts = {"个人": 0, "同桌": 0, "分组": 0}
    for lesson in lessons:
        counts[lesson["form"]] += 1
    summary = (
        f"任务形式合计：个人 {counts['个人']} 节 · 同桌 {counts['同桌']} 节 · 分组 {counts['分组']} 节。"
        "点击课时编号可跳到完整课程表对应行。模块、天数和上方路径是整理归纳，过关标准与产出用语来自提纲原文。"
    )
    write_merged(
        ws, row, 1, 6, summary,
        font(8, color=MUTED), fill(WHITE),
        Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1), 32,
    )
    apply_page(ws, "1:2", fit_height=0, zoom=120)
    ws.freeze_panes = "A4"
    ws.print_area = f"A1:F{row}"
    ws.sheet_properties.tabColor = "163A47"
    ws.oddHeader.right.text = "三日地图"
    return ws


def box(ws, r1, c1, r2, c2, color):
    medium = side("medium", color)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            current = cell.border
            cell.border = Border(
                left=medium if c == c1 else current.left,
                right=medium if c == c2 else current.right,
                top=medium if r == r1 else current.top,
                bottom=medium if r == r2 else current.bottom,
            )


def verify(path: Path, lessons):
    wb = load_workbook(path)
    expected = ["三日地图", "完整课程表", "产出与考核", "原文八列"]
    if wb.sheetnames != expected:
        raise SystemExit(f"工作表顺序不符：{wb.sheetnames}")
    full = wb["完整课程表"]
    source = wb["原文八列"]
    check = wb["产出与考核"]
    for offset, lesson in enumerate(lessons):
        row = 4 + offset
        full_values = [full.cell(row, col).value for col in range(3, 11)]
        source_values = [source.cell(row, col).value for col in range(1, 9)]
        original = [
            lesson["id"], lesson["topic"], lesson["objective"], lesson["content"],
            lesson["method"], lesson["task"], lesson["output"], lesson["standard"],
        ]
        if full_values != original or source_values != original:
            raise SystemExit(f"{lesson['id']} 与原文不一致")
        if check.cell(row, 5).value != lesson["output"]:
            raise SystemExit(f"{lesson['id']} 产出未进入考核表")
        if check.cell(row, 6).value != lesson["standard"]:
            raise SystemExit(f"{lesson['id']} 考核标准未进入考核表")
        if full.cell(row, 11).value != lesson["form"]:
            raise SystemExit(f"{lesson['id']} 任务形式不符")
    if check["H2"].value != '=COUNTIF(G4:G27,"已完成")&" / 24 节"':
        raise SystemExit("完成进度公式丢失")
    if "gork" not in full["F8"].value:
        raise SystemExit("L5 模型名称被改写")
    print(f"已写入 {path}")
    print(f"工作表：{'、'.join(wb.sheetnames)}")
    print(f"课时 {len(lessons)} · 原文八列逐格对照通过")


def main():
    lessons = load_lessons()
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    _, locations = build_full(wb, lessons)
    build_check(wb, lessons)
    build_source(wb, lessons)
    build_map(wb, lessons, locations)
    wb.properties.title = "三天实训课程表"
    wb.properties.subject = "24 课时课程表"
    wb.properties.creator = "课程表整理"
    wb.properties.description = (
        "依据授课提纲整理的三天 24 课时课程表，含三日地图、完整课程表、产出与考核、原文八列。"
    )
    wb.properties.keywords = "课程表,Codex,dsh,短剧,广告"
    wb.active = 0
    wb.save(OUT)
    verify(OUT, lessons)


if __name__ == "__main__":
    main()
