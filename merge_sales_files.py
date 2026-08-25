# -*- coding: utf-8 -*-
"""合并并清洗多个 Excel/CSV 销售文件。"""

from pathlib import Path
import logging

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_FILE = OUTPUT_DIR / "merged_sales_report.xlsx"
PDF_DIR = OUTPUT_DIR / "pdf"
PDF_FILE = PDF_DIR / "daily_sales_report.pdf"
LOG_FILE = OUTPUT_DIR / "process.log"

REQUIRED_COLUMNS = {"日期", "销售员", "产品", "数量", "金额"}


def setup_logging() -> None:
    """同时把信息输出到屏幕和日志文件。"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def read_one_file(file_path: Path) -> pd.DataFrame | None:
    """读取并检查一个 Excel 或 CSV 文件。"""
    try:
        if file_path.suffix.lower() == ".csv":
            data = pd.read_csv(file_path)
        elif file_path.suffix.lower() == ".json":
            data = pd.read_json(file_path)
        else:
            data = pd.read_excel(file_path)
    except Exception as error:
        logging.error("读取失败：%s；原因：%s", file_path.name, error)
        return None

    if data.empty:
        logging.warning("跳过空文件：%s", file_path.name)
        return None

    missing_columns = REQUIRED_COLUMNS - set(data.columns)
    if missing_columns:
        logging.error(
            "跳过文件：%s；缺少列：%s",
            file_path.name,
            ", ".join(sorted(missing_columns)),
        )
        return None

    data["来源文件"] = file_path.name
    logging.info("读取成功：%s，共 %d 行", file_path.name, len(data))
    return data


def merge_files() -> pd.DataFrame:
    """读取 input 中的 Excel/CSV 文件并清洗合并。"""
    files = sorted(
        list(INPUT_DIR.glob("*.xlsx"))
        + list(INPUT_DIR.glob("*.csv"))
        + list(INPUT_DIR.glob("*.json"))
    )

    valid_data = []
    for file_path in files:
        if file_path.name.startswith("~$"):
            continue
        data = read_one_file(file_path)
        if data is not None:
            valid_data.append(data)

    if not valid_data:
        raise ValueError("没有可合并的有效文件")

    merged = pd.concat(valid_data, ignore_index=True)

    before_count = len(merged)
    # 来源文件不同，也应把相同的销售记录视为重复数据。
    dedup_columns = ["日期", "销售员", "产品", "数量", "金额"]
    merged = merged.drop_duplicates(subset=dedup_columns)
    logging.info("去重完成，删除 %d 行重复数据", before_count - len(merged))

    merged["日期"] = pd.to_datetime(merged["日期"], errors="coerce")
    merged["数量"] = pd.to_numeric(merged["数量"], errors="coerce")
    merged["金额"] = pd.to_numeric(merged["金额"], errors="coerce")

    invalid_quantity = merged["数量"].isna().sum()
    invalid_amount = merged["金额"].isna().sum()
    if invalid_quantity > 0:
        logging.warning("发现 %d 条数量无效", invalid_quantity)
    if invalid_amount > 0:
        logging.warning("发现 %d 条金额无效", invalid_amount)

    before_clean = len(merged)
    merged = merged.dropna(
        subset=["日期", "销售员", "产品", "数量", "金额"]
    )
    logging.info("清理缺失数据，删除 %d 行", before_clean - len(merged))

    invalid_quantity_value = (merged["数量"] <= 0).sum()
    invalid_amount_value = (merged["金额"] <= 0).sum()
    if invalid_quantity_value > 0:
        logging.warning("发现 %d 条数量不合理", invalid_quantity_value)
    if invalid_amount_value > 0:
        logging.warning("发现 %d 条金额不合理", invalid_amount_value)

    before_reasonable = len(merged)
    merged = merged[(merged["数量"] > 0) & (merged["金额"] > 0)]
    logging.info(
        "清理不合理数据，删除 %d 行",
        before_reasonable - len(merged),
    )
    return merged


def save_report(data: pd.DataFrame) -> None:
    """保存明细和汇总工作表。"""
    summary = pd.DataFrame(
        {
            "指标": ["总行数", "总数量", "总金额", "平均金额"],
            "数值": [
                len(data),
                data["数量"].sum(),
                data["金额"].sum(),
                data["金额"].mean(),
            ],
        }
    )

    daily_summary = (
        data.assign(日期=data["日期"].dt.strftime("%Y-%m-%d"))
        .groupby("日期", as_index=False)
        .agg(
            记录数=("产品", "count"),
            总数量=("数量", "sum"),
            总金额=("金额", "sum"),
        )
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="合并明细")
        summary.to_excel(writer, index=False, sheet_name="汇总")
        daily_summary.to_excel(writer, index=False, sheet_name="按日汇总")

    logging.info("报表已保存：%s", OUTPUT_FILE)


def save_pdf_report(data: pd.DataFrame) -> None:
    """把按日汇总结果导出为 PDF。"""
    PDF_DIR.mkdir(exist_ok=True)
    font_path = r"C:\Windows\Fonts\msyh.ttc"
    pdfmetrics.registerFont(TTFont("MicrosoftYaHei", font_path, subfontIndex=0))

    daily_summary = (
        data.assign(日期=data["日期"].dt.strftime("%Y-%m-%d"))
        .groupby("日期", as_index=False)
        .agg(
            记录数=("产品", "count"),
            总数量=("数量", "sum"),
            总金额=("金额", "sum"),
        )
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ChineseTitle",
        parent=styles["Title"],
        fontName="MicrosoftYaHei",
        fontSize=18,
        alignment=TA_CENTER,
        spaceAfter=8 * mm,
    )
    body_style = ParagraphStyle(
        "ChineseBody",
        parent=styles["BodyText"],
        fontName="MicrosoftYaHei",
        fontSize=10,
        leading=16,
    )

    table_data = [["日期", "记录数", "总数量", "总金额"]]
    for _, row in daily_summary.iterrows():
        table_data.append(
            [
                str(row["日期"]),
                str(int(row["记录数"])),
                str(int(row["总数量"])),
                f"{row['总金额']:.2f}",
            ]
        )

    document = SimpleDocTemplate(
        str(PDF_FILE),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
    )
    table = Table(table_data, colWidths=[45 * mm, 30 * mm, 30 * mm, 40 * mm])
    table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "MicrosoftYaHei"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2563EB")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9E2F3")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F8FC")]),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story = [
        Paragraph("销售日报", title_style),
        Paragraph("按日期汇总销售记录、数量和金额。", body_style),
        Spacer(1, 6 * mm),
        table,
    ]
    document.build(story)
    logging.info("PDF 日报已保存：%s", PDF_FILE)


def main() -> None:
    setup_logging()
    try:
        merged_data = merge_files()
        save_report(merged_data)
        save_pdf_report(merged_data)
        logging.info("处理完成，共输出 %d 行", len(merged_data))
    except Exception as error:
        logging.exception("处理失败：%s", error)


if __name__ == "__main__":
    main()
