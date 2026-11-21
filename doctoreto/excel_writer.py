"""Excel output adapter; keeps the original six article columns."""

import os
import tempfile
from collections import Counter
from pathlib import Path

import jdatetime
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font

from .aggregate import monthly_categories
from .models import ArticleRecord


CELL_LIMIT = 32767
PERSIAN_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")


def _persian_date(record: ArticleRecord) -> str | None:
    if record.published_date is None:
        return None
    value = jdatetime.date.fromgregorian(date=record.published_date)
    return value.strftime("%Y/%m/%d").translate(PERSIAN_DIGITS)


def _header(sheet, values: list[str]) -> None:
    sheet.sheet_view.rightToLeft = True
    sheet.append(values)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"


def _finish_table(sheet) -> None:
    sheet.auto_filter.ref = sheet.dimensions


def _monthly_sheet(workbook: Workbook, key: str, total: int, topics: Counter[str]) -> None:
    persian_month = key.translate(PERSIAN_DIGITS)
    sheet = workbook.create_sheet(f"ماه-{persian_month}")
    _header(sheet, ["موضوع", "تعداد مقالات", "درصد از مقالات همان ماه"])
    for topic, count in topics.most_common():
        sheet.append([topic, count, count / total])
        sheet.cell(sheet.max_row, 3).number_format = "0.0%"
    sheet["E1"] = f"تعداد مقالات این ماه: {total}"
    sheet["E2"] = "مقاله‌ها ممکن است چند دسته‌بندی داشته باشند."
    sheet["E3"] = "بنابراین درصدها الزاماً جمع ۱۰۰٪ نمی‌شوند."
    for column, width in {"A": 45, "B": 18, "C": 28}.items():
        sheet.column_dimensions[column].width = width
    sheet.auto_filter.ref = f"A1:C{sheet.max_row}"

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.overlap = 0
    chart.style = 10
    chart.title = f"درصد مقالات بر اساس موضوع - {persian_month}"
    chart.x_axis.title = "موضوع"
    chart.y_axis.title = "درصد مقالات"
    chart.y_axis.numFmt = "0%"
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 1
    chart.height = 15
    chart.width = 28
    chart.legend = None
    chart.add_data(Reference(sheet, min_col=3, min_row=1, max_row=sheet.max_row), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=sheet.max_row))
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showCatName = True
    chart.dataLabels.showVal = False
    chart.dataLabels.showSerName = False
    chart.dataLabels.dLblPos = "outEnd"
    sheet.add_chart(chart, "E5")


def build_workbook(records: list[ArticleRecord]) -> Workbook:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "مقالات"
    _header(sheet, ["تاریخ انتشار", "عنوان", "دسته‌بندی", "برچسب‌ها", "متن مقاله HTML", "لینک", "وضعیت HTML"])
    chunks: list[list[object]] = []
    errors: list[list[str]] = []

    for record in records:
        html = record.html
        html_status = "موجود" if html else "ناموجود"
        if html and len(html) > CELL_LIMIT:
            html_status = "در شیت بخش‌ها"
            for number, start in enumerate(range(0, len(html), CELL_LIMIT), 1):
                chunks.append([record.url, record.title, number, html[start:start + CELL_LIMIT]])
            html = None
        sheet.append([
            _persian_date(record), record.title, "، ".join(record.categories) or None,
            "، ".join(record.tags) or None, html, record.url, html_status,
        ])
        if record.issues:
            errors.append([record.url, record.status, "; ".join(record.issues)])

    for column, width in {"A": 18, "B": 60, "C": 30, "D": 40, "E": 100, "F": 80, "G": 20}.items():
        sheet.column_dimensions[column].width = width
    _finish_table(sheet)

    if chunks:
        chunk_sheet = workbook.create_sheet("HTMLهای طولانی - بخش‌ها")
        _header(chunk_sheet, ["لینک مقاله", "عنوان", "شماره بخش", "بخش HTML"])
        for row in chunks:
            chunk_sheet.append(row)
        for column, width in {"A": 80, "B": 60, "C": 14, "D": 100}.items():
            chunk_sheet.column_dimensions[column].width = width
        _finish_table(chunk_sheet)

    if errors:
        error_sheet = workbook.create_sheet("موارد نیازمند بررسی")
        _header(error_sheet, ["لینک مقاله", "وضعیت", "خطا"])
        for row in errors:
            error_sheet.append(row)
        for column, width in {"A": 80, "B": 18, "C": 80}.items():
            error_sheet.column_dimensions[column].width = width
        _finish_table(error_sheet)

    for key, (total, topics) in monthly_categories(records).items():
        _monthly_sheet(workbook, key, total, topics)
    return workbook


def save_workbook(records: list[ArticleRecord], path: Path, overwrite: bool = False) -> None:
    """Write beside the destination, then replace only after a complete save."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and not overwrite:
        raise FileExistsError(f"Output already exists: {path}")

    workbook = build_workbook(records)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(prefix=".doctoreto-", suffix=".xlsx", dir=path.parent, delete=False) as handle:
            temporary_path = Path(handle.name)
        workbook.save(temporary_path)
        if path.exists() and not overwrite:
            raise FileExistsError(f"Output already exists: {path}")
        os.replace(temporary_path, path)
    finally:
        workbook.close()
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
