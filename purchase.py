"""Расчёт стоимости заявки по прайсам поставщиков.

Строки двух файлов связывает артикул. Товары могут идти в разном порядке.
"""
import argparse
import csv
import re
from contextlib import closing
from decimal import Decimal
from pathlib import Path
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from allocation import Offer, build_quote


def input_rows(path):
    """Возвращает номер строки и значения первого листа Excel или CSV."""
    if path.suffix.lower() == ".xlsx":
        try:
            from openpyxl import load_workbook
        except ModuleNotFoundError as error:
            raise ValueError("для Excel установите зависимости: python -m pip install -r requirements.txt") from error
        try:
            with path.open("rb") as source, closing(load_workbook(source, read_only=True)) as book:
                if not book.worksheets:
                    raise ValueError(f"{path}: нет листа с данными")
                for number, row in enumerate(book.worksheets[0].values, 1):
                    yield number, ["" if value is None else str(value).strip() for value in row]
        except (BadZipFile, KeyError, ParseError) as error:
            raise ValueError(f"{path}: повреждён файл Excel") from error
        return

    if path.suffix.lower() != ".csv":
        raise ValueError(f"{path}: нужен файл .csv или .xlsx")
    with path.open(encoding="utf-8-sig", newline="") as source:
        reader = csv.reader(source, delimiter=";", strict=True)
        try:
            for row in reader:
                yield reader.line_num, [value.strip() for value in row]
        except csv.Error as error:
            raise ValueError(f"{path}, строка {reader.line_num}: {error}") from error


def read_table(path, required, parse_row, aliases=None):
    items = {}
    with closing(input_rows(path)) as rows:
        _, headers = next(rows, (1, []))
        if aliases:
            headers = [aliases.get(name, name) for name in headers]
        names = [name for name in headers if name]
        if len(names) != len(set(names)) or not {"артикул", required}.issubset(names):
            raise ValueError(f"{path}, строка 1: нужны уникальные столбцы «артикул» и «{required}»")
        for number, row in rows:
            if not any(row):
                continue
            location = f"{path}, строка {number}"
            if len(row) != len(headers):
                raise ValueError(f"{location}: проверьте число полей")
            # Начальные нули важны: у поставщика «001» и «1» могут означать разные товары.
            record = dict(zip(headers, row))
            sku = record["артикул"]
            if not sku or sku in items:
                raise ValueError(f"{location}: пустой или повторяющийся артикул «{sku}»")
            try:
                items[sku] = parse_row(record)
            except ValueError as error:
                raise ValueError(f"{location}: {error}") from error
    if not items:
        raise ValueError(f"{path}: нет товарных позиций")
    return items


def read_items(path, field, parse):
    return read_table(path, field, lambda row: parse(row[field]))


def read_offers(path):
    def parse_offer(row):
        stock = None
        if "остаток, шт" in row:
            if not re.fullmatch(r"[0-9]+", row["остаток, шт"]):
                raise ValueError("остаток должен быть целым неотрицательным числом штук")
            stock = int(row["остаток, шт"])
        pack = row.get("штук в упаковке", "1")
        if not re.fullmatch(r"[0-9]+", pack) or int(pack) < 1:
            raise ValueError("число штук в упаковке должно быть целым положительным")
        return Offer(path.stem, price(row["цена"]), int(pack), stock)

    return read_table(path, "цена", parse_offer, {"цена, ₽/шт": "цена"})


def quantity(value: str) -> int:
    """Количество в заявке измеряется в целых штуках."""
    if not re.fullmatch(r"[0-9]+", value) or int(value) < 1:
        raise ValueError("количество должно быть целым положительным числом")
    return int(value)


def price(value: str) -> Decimal:
    """Принимает цену с точкой или запятой. Decimal сохраняет копейки при расчёте."""
    if not re.fullmatch(r"[0-9]+([.,][0-9]{1,2})?", value):
        raise ValueError("нужна неотрицательная цена с точностью до копеек")
    return Decimal(value.replace(",", "."))


def rubles(value: Decimal) -> str:
    return f"{value:.2f}".replace(".", ",") + " ₽"


def report(request: dict[str, int], prices: dict[str, Decimal]) -> str:
    """Считает стоимость строк в порядке заявки и подписывает итог, если каких-то цен нет."""
    lines = []
    total = Decimal(0)
    missing = 0
    for sku, count in request.items():
        # Проверяем наличие артикула: нулевая цена в прайсе тоже действительна.
        if sku not in prices:
            missing += 1
            lines.append(f"{sku}: {count} шт. — нет в прайсе")
            continue
        amount = count * prices[sku]
        total += amount
        lines.append(f"{sku}: {count} × {rubles(prices[sku])} = {rubles(amount)}")
    label = f"По найденным позициям (без цены: {missing})" if missing else "Итого"
    lines.append(f"{label}: {rubles(total)}")
    return "\n".join(lines)


def supplier_report(quote):
    lines = []
    total = Decimal(0)
    incomplete = 0
    for item in quote:
        incomplete += bool(item.shortage)
        lines.append(f"{item.sku}: заявка {item.demand}, заказать {item.units}, "
                     f"излишек {item.surplus}, дефицит {item.shortage} шт.")
        if not item.listed:
            lines.append("  Нет в прайсах")
        for order in item.orders:
            lines.append(f"  {order.offer.supplier}: {order.packs} уп. × {order.offer.pack} шт., "
                         f"{rubles(order.offer.price)}/шт. — {rubles(order.amount)}")
            total += order.amount
    label = f"Неполный заказ (позиций с дефицитом: {incomplete})" if incomplete else "Итого"
    return "\n".join([*lines, f"{label}: {rubles(total)}"])


def main():
    parser = argparse.ArgumentParser(description="Расчёт заявки по прайсам поставщиков")
    parser.add_argument("request", type=Path, help="Заявка (.csv или .xlsx)")
    parser.add_argument("prices", type=Path, nargs="+", help="Прайсы поставщиков (.csv или .xlsx)")
    parser.add_argument("-o", "--output", type=Path, help="Сохранить отчёт в .xlsx")
    args = parser.parse_args()
    try:
        request = read_items(args.request, "количество", quantity)
        names = [path.stem.casefold() for path in args.prices]
        if len(set(names)) != len(names):
            raise ValueError("дайте прайсам разных поставщиков разные имена файлов")
        catalogs = [read_offers(path) for path in args.prices]
        quote = build_quote(request, catalogs)
        if args.output:
            if args.output.suffix.lower() != ".xlsx":
                raise ValueError("для отчёта укажите файл с расширением .xlsx")
            if args.output.resolve() in {p.resolve() for p in [args.request, *args.prices]}:
                raise ValueError("путь отчёта совпадает с входным файлом")
            try:
                from excel_report import save_report
            except ModuleNotFoundError as error:
                raise ValueError("для Excel установите зависимости: python -m pip install -r requirements.txt") from error
            save_report(args.output, quote)
    except (ValueError, OSError, csv.Error) as error:
        parser.exit(2, f"Ошибка: {error}\n")
    if len(catalogs) == 1 and all(o.pack == 1 and o.stock is None for o in catalogs[0].values()):
        print(report(request, {sku: offer.price for sku, offer in catalogs[0].items()}))
    else:
        print(supplier_report(quote))
    if args.output:
        print(f"Отчёт сохранён: {args.output}")


if __name__ == "__main__":
    main()
