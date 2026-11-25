#!/usr/bin/env python3
"""Validate Guizhou recruiting CSV files and rebuild README plus formatted Excel."""
from __future__ import annotations

import csv
import re
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data" / "jobs.csv"
WATCH = ROOT / "data" / "watchlist_and_excluded.csv"
FIELDS = [
    "id", "unit", "unit_type", "recruitment_type", "target", "education",
    "match_tags", "positions", "work_locations", "application", "deadline",
    "deadline_type", "status", "announcement_date", "source_url",
    "source_level", "last_verified", "notes",
]
HEADERS = [
    "序号", "单位名称", "单位类型", "招聘类型", "招聘对象", "招聘学历",
    "匹配方向/专业代码", "相关岗位", "岗位城市/地点", "网申链接/投递方式",
    "招聘截止时间", "截止口径", "当前状态", "公告/发布日期", "信息来源",
    "来源级别", "最近核验", "备注",
]
WIDTHS = [7, 34, 20, 26, 16, 30, 34, 42, 38, 48, 24, 34, 32, 20, 50, 28, 15, 62]
WATCH_FIELDS = ["unit", "related_position", "category", "deadline", "reason", "source_url", "last_verified"]


def load(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        rows = list(reader)
        if reader.fieldnames is None:
            raise ValueError(f"{path}: 缺少表头")
        if any(None in row for row in rows):
            raise ValueError(f"{path}: 存在列数错误")
        return rows


def validate(rows: list[dict[str, str]], watch: list[dict[str, str]]) -> None:
    if not rows or list(rows[0]) != FIELDS:
        raise ValueError("主表为空或字段顺序不符合18字段约定")
    if len({row["id"] for row in rows}) != len(rows) or len({row["unit"] for row in rows}) != len(rows):
        raise ValueError("序号或单位名称重复")
    for row in rows:
        for key in ("unit", "match_tags", "positions", "work_locations", "deadline", "status", "source_url", "last_verified"):
            if not row[key].strip():
                raise ValueError(f"{row['unit']}: {key}为空")
        if not row["source_url"].startswith(("http://", "https://")):
            raise ValueError(f"{row['unit']}: 来源URL无效")
        date.fromisoformat(row["last_verified"])
    if watch and list(watch[0]) != WATCH_FIELDS:
        raise ValueError("待关注表字段顺序不符合7字段约定")


def is_expired(row: dict[str, str]) -> bool:
    return row["status"].startswith("已截止")


def deadline_key(row: dict[str, str]) -> tuple:
    match = re.match(r"(20\d{2}-\d{2}-\d{2})", row["deadline"])
    if is_expired(row):
        return (3, match.group(1) if match else "", row["unit"])
    if match:
        return (0, match.group(1), row["unit"])
    if "招满即止" in row["deadline"]:
        return (1, "9999-12-31", row["unit"])
    return (2, "9999-12-31", row["unit"])


def md(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", "<br>")


def build_readme(rows: list[dict[str, str]], watch: list[dict[str, str]]) -> None:
    ordered = sorted(rows, key=deadline_key)
    active = [row for row in rows if not is_expired(row)]
    expired = [row for row in rows if is_expired(row)]
    as_of = max(row["last_verified"] for row in rows)
    insar = sum("insar" in (row["match_tags"] + row["positions"]).lower() for row in active)
    code = sum("085704" in row["match_tags"] for row in rows)
    unknown = sum("未注明" in row["deadline"] for row in active)
    lines = [
        "# 2027届贵州 InSAR / 测绘工程 / 资源与环境秋招岗位表", "",
        f"> 最近核验：{as_of}（北京时间）｜主表 **{len(rows)} 家**：仍可投/岗位页可查 **{len(active)} 家**，已截止 **{len(expired)} 家**｜每天 12:00 更新", "",
        "面向 **InSAR、测绘工程（按专业代码 085704 整理）、资源与环境（0857，限测绘/遥感/地理信息相关方向）**，且岗位工作地点明确涉及贵州省内的公开招聘信息。", "",
        "## 数据概览", "",
        f"- 当前仍可投、公告或官方岗位页仍可查：**{len(active)} 家**；已截止但专业、地点匹配：**{len(expired)} 家**，按用户要求保留在主表末尾。",
        f"- 当前在招记录中明确写出 **InSAR**：**{insar} 家**；主表明确标注 **085704**：**{code} 家**（含已截止记录）。",
        f"- 在招记录中未注明固定截止：**{unknown} 家**，不代表长期有效或名额充足。",
        f"- 另有待确认、专业不匹配或未找到本轮公告：**{len(watch)} 条**，见 [`data/watchlist_and_excluded.csv`](data/watchlist_and_excluded.csv)。",
        "- 完整机器可读数据：[`data/jobs.csv`](data/jobs.csv)；本地运行生成脚本可得到 `2027_guizhou_campus_jobs.xlsx`。", "",
        "## 贵州岗位汇总（仍在招优先，已截止置底）", "",
        "| 单位 | 当前状态 | 匹配方向/代码 | 相关岗位 | 贵州工作地点 | 截止时间 | 来源 |",
        "|---|---|---|---|---|---|---|",
    ]
    for row in ordered:
        lines.append(
            f"| {md(row['unit'])} | {md(row['status'])} | {md(row['match_tags'])} | "
            f"{md(row['positions'])} | {md(row['work_locations'])} | {md(row['deadline'])} | "
            f"[公告/职位页]({row['source_url']}) |"
        )
    lines += [
        "", "## 截止时间与状态口径", "",
        "- **公告明确**：按单位公告或高校就业网转载公告原文记录。",
        "- **官方招聘系统未注明**：官方职位仍展示，但没有可核验的固定截止日期。",
        "- **高校校内内推截止**：只代表该校内推通道截止，不自动等同于企业所有招聘通道关闭。",
        "- **未注明**：应尽快投递并联系招聘单位确认，不能理解为长期有效。",
        "- **已截止置底**：符合专业和贵州地点口径、但已发现截止时间的记录仍保留，便于复盘和关注补录/新批次。", "",
        "## 收录边界", "",
        "1. 地域按**实际工作地点在贵州省内**筛选；仅总部在贵州、拥有测绘资质或全国项目可能分配贵州，不足以单独纳入。",
        "2. 用户原文“测绘工程 087504”暂按常见专业代码 **085704** 整理；0857只收录测绘、遥感、GIS、InSAR等相关方向。",
        "3. 同一集团与独立下属单位在岗位、地点或投递入口不同时分别列出，可能存在计划重叠，投递时需注明意向单位。",
        "4. 宣讲会结束不等于网申截止；高校页面“已过期”也不自动等于集团全部招聘结束，需结合正式公告判断。",
        "5. “所有单位”无法由公开网络一次性证明穷尽。本仓库是**公开渠道可核验合集**，通过每日更新持续补漏，不作绝对穷尽承诺。", "",
        "## 每日更新规则", "",
        "- 每天 12:00（Asia/Shanghai）检查新增公告、网申状态、专业目录、贵州工作地点及截止时间。",
        "- 新增记录必须有来源；发现已截止的岗位保留在 `data/jobs.csv` 主表末尾，并更新状态、截止口径和核验日期。",
        "- 更新 CSV 后运行 `python3 scripts/build_outputs.py`，重新生成 README 与 Excel。",
        "- GitHub Actions 在推送及每天 04:00 UTC（北京时间 12:00）执行数据校验和生成物同步检查；实际检索与内容更新由每日任务执行。", "",
        "## 本地重建", "", "```bash", "python3 -m pip install -r requirements.txt",
        "python3 scripts/build_outputs.py", "```", "", "## 免责声明", "",
        "招聘状态、名额和截止时间可能随时调整。本仓库仅做信息整理，不代替招聘单位正式公告，不构成录用承诺。投递前请再次核对官方公告与岗位状态。", "",
    ]
    (ROOT / "README.md").write_text("\n".join(lines), encoding="utf-8")


def style_sheet(ws, widths: list[int]) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    ws.sheet_view.showGridLines = False
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color="B8C7D9"))
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[row[0].row].height = 66


def add_first_url_hyperlink(cell) -> None:
    match = re.search(r"https?://[^；;\s]+", str(cell.value or ""))
    if match:
        cell.hyperlink = match.group(0)
        cell.font = Font(name="微软雅黑", size=9, color="0563C1", underline="single")


def build_excel(rows: list[dict[str, str]], watch: list[dict[str, str]]) -> None:
    as_of = max(date.fromisoformat(row["last_verified"]) for row in rows)
    active_count = sum(not is_expired(row) for row in rows)
    expired_count = len(rows) - active_count
    ordered = sorted(rows, key=deadline_key)
    wb = Workbook()
    ws = wb.active
    ws.title = "贵州岗位汇总"
    ws.append(HEADERS)
    for row in ordered:
        ws.append([row[key] for key in FIELDS])
    style_sheet(ws, WIDTHS)
    for col in (10, 15):
        for row_index in range(2, ws.max_row + 1):
            add_first_url_hyperlink(ws.cell(row_index, col))
    for row_index, row in enumerate(ordered, 2):
        status_cell = ws.cell(row_index, 13)
        if is_expired(row):
            for cell in ws[row_index]:
                cell.fill = PatternFill("solid", fgColor="F2F2F2")
                cell.font = Font(name="微软雅黑", size=10, color="666666")
            status_cell.fill = PatternFill("solid", fgColor="FCE8E6")
            status_cell.font = Font(name="微软雅黑", size=10, bold=True, color="B42318")
            ws.cell(row_index, 11).font = Font(name="微软雅黑", size=10, bold=True, color="B42318")
            continue
        status_cell.fill = PatternFill("solid", fgColor="E8F3E8")
        status_cell.font = Font(name="微软雅黑", size=10, bold=True, color="26734D")
        match = re.match(r"(20\d{2}-\d{2}-\d{2})", row["deadline"])
        if match and 0 <= (date.fromisoformat(match.group(1)) - as_of).days <= 7:
            ws.cell(row_index, 11).fill = PatternFill("solid", fgColor="FCE8E6")
            ws.cell(row_index, 11).font = Font(name="微软雅黑", size=10, bold=True, color="B42318")

    info = wb.create_sheet("使用说明")
    info_rows = [
        ("项目", "说明"), ("最近核验", f"{as_of.isoformat()}（北京时间）"),
        ("记录数", f"主表{len(rows)}家：仍在招/可查{active_count}家，已截止{expired_count}家；另{len(watch)}条待关注或排除"),
        ("排序", "仍在招/官方职位页可查在前，已截止统一置底；明确截止日期优先"),
        ("地域口径", "岗位实际工作地点明确涉及贵州省内，不仅看单位总部或测绘资质"),
        ("专业代码", "用户原文087504暂按085704测绘工程整理；0857限测绘、遥感、GIS、InSAR相关方向"),
        ("颜色", "绿色状态=在招/可查；红色=已截止或临近截止；灰色行=已截止保留记录"),
        ("完整性", "公开渠道可核验合集，不作绝对穷尽承诺；每天12点持续更新"),
        ("提示", "未注明截止不等于长期有效；宣讲会结束不等于网申关闭；投递前请复核官方公告"),
    ]
    for row in info_rows:
        info.append(row)
    style_sheet(info, [24, 115])

    watch_ws = wb.create_sheet("待关注与已排除")
    watch_ws.append(["单位", "相关岗位", "分类", "截止/状态", "说明", "来源", "最近核验"])
    for row in watch:
        watch_ws.append([row[key] for key in WATCH_FIELDS])
    style_sheet(watch_ws, [38, 38, 25, 32, 78, 55, 15])
    for row_index in range(2, watch_ws.max_row + 1):
        add_first_url_hyperlink(watch_ws.cell(row_index, 6))

    output = ROOT / "2027_guizhou_campus_jobs.xlsx"
    wb.save(output)
    check = load_workbook(output, read_only=True)
    assert check.sheetnames == ["贵州岗位汇总", "使用说明", "待关注与已排除"]
    assert check["贵州岗位汇总"].max_row == len(rows) + 1
    check.close()


def main() -> None:
    rows, watch = load(DATA), load(WATCH)
    validate(rows, watch)
    build_readme(rows, watch)
    build_excel(rows, watch)
    active = sum(not is_expired(row) for row in rows)
    print(f"OK: {active} active/checkable records; {len(rows) - active} expired records at bottom; {len(watch)} watch/excluded records")


if __name__ == "__main__":
    main()
