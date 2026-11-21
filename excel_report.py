"""Excel-отчёт по рассчитанной закупке."""
import re
from collections import defaultdict
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


MONEY = '#,##0.00 "₽"'
SUMMARY = ["Артикул", "Запрошено, шт", "Заказать, шт", "Излишек, шт", "Дефицит, шт",
           "Сумма, ₽", "Статус"]
ORDER = ["Артикул", "Запрошено всего, шт", "Упаковок", "Штук в упаковке", "Заказать, шт",
         "Цена за штуку, ₽", "Сумма, ₽"]


def write_row(sheet, number, values):
    # Артикулы и имена поставщиков — текст, в том числе значения вида «=001».
    for column, value in enumerate(values, 1):
        cell = sheet.cell(number, column, value)
        if isinstance(value, str):
            cell.data_type = "s"


def write_table(sheet, title, amount, status, headers, rows, money_columns):
    write_row(sheet, 1, [title])
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    sheet["A1"].font = Font(size=16, bold=True, color="164E58")
    sheet.row_dimensions[1].height = 28
    write_row(sheet, 2, ["Сумма заказа", amount, status])
    sheet["B2"].number_format = MONEY
    sheet["B2"].font = Font(bold=True, color="164E58")
    sheet.merge_cells(start_row=2, start_column=3, end_row=2, end_column=len(headers))
    sheet["C2"].font = Font(bold=True, color="884718" if "дефицит" in status.lower() else "164E58")
    write_row(sheet, 3, headers)
    for cell in sheet[3]:
        cell.fill = PatternFill("solid", fgColor="205765")
        cell.font = Font(bold=True, color="FFFFFF")
    widths = [max(len(name) + 3, 16) for name in headers]
    for number, values in enumerate(rows, 4):
        write_row(sheet, number, values)
        for index, value in enumerate(values):
            widths[index] = max(widths[index], len(str(value)) + 3)
        for column in money_columns:
            sheet.cell(number, column).number_format = MONEY
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = min(width, 65)
    sheet["C2"].alignment = Alignment(horizontal="left")
    sheet.freeze_panes = "B4"
    sheet.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{sheet.max_row}"
    sheet.sheet_view.showGridLines = False


def supplier_sheet_name(supplier, used):
    base = re.sub(r"[\\/*?:\[\]]", "_", supplier).strip("'") or "Поставщик"
    name = base[:31].rstrip("'")
    index = 1
    while name.casefold() in used:
        index += 1
        suffix = f" ({index})"
        name = base[:31 - len(suffix)] + suffix
    return name


def save_report(path, quote):
    book = Workbook()
    summary = book.active
    summary.title = "Сводка"
    attention = book.create_sheet("Требуют внимания")
    rows, issues = [], []
    suppliers = defaultdict(list)
    shortage_count = 0
    for item in quote:
        status = "Покрыто"
        if item.shortage:
            shortage_count += 1
            status = "Недостаточно полных упаковок" if item.listed else "Нет в прайсах"
        elif item.surplus:
            status = "Излишек из-за упаковки"
        row = [item.sku, item.demand, item.units, item.surplus, item.shortage, item.amount, status]
        rows.append(row)
        if item.shortage or item.surplus:
            issues.append(row)
        for order in item.orders:
            suppliers[order.offer.supplier].append([
                item.sku, item.demand, order.packs, order.offer.pack, order.units,
                order.offer.price, order.amount,
            ])
    status = f"Неполный заказ · позиций с дефицитом: {shortage_count}" if shortage_count else "Заявка покрыта"
    write_table(summary, "Сводка закупки", sum(item.amount for item in quote), status, SUMMARY, rows, [6])
    write_table(attention, "Позиции для проверки", sum(row[5] for row in issues),
                f"Позиций для проверки: {len(issues)}", SUMMARY, issues, [6])
    used = {name.casefold() for name in [*book.sheetnames, "History"]}
    for supplier, orders in suppliers.items():
        name = supplier_sheet_name(supplier, used)
        used.add(name.casefold())
        sheet = book.create_sheet(name)
        write_table(sheet, supplier, sum(row[6] for row in orders), "Заказ поставщику", ORDER, orders, [6, 7])

    # Готовый файл заменяет предыдущий отчёт только после успешной записи всей книги.
    with NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as temp:
        temporary = Path(temp.name)
    try:
        book.save(temporary)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
        book.close()
