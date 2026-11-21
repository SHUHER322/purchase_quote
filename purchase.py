"""Расчёт стоимости заявки по прайсу одного поставщика.

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


def read_items(path, field, parse):
    """Читает пары «артикул — значение». Для заявки разбираем количество, для прайса — цену."""
    items = {}
    with closing(input_rows(path)) as rows:
        _, headers = next(rows, (1, []))
        if headers.count("артикул") != 1 or headers.count(field) != 1:
            raise ValueError(f"{path}, строка 1: нужны столбцы «артикул» и «{field}»")
        sku_column = headers.index("артикул")
        value_column = headers.index(field)
        for number, row in rows:
            if not any(row):
                continue
            location = f"{path}, строка {number}"
            if len(row) != len(headers):
                raise ValueError(f"{location}: проверьте число полей")
            # Начальные нули важны: у поставщика «001» и «1» могут означать разные товары.
            sku = row[sku_column]
            if not sku or sku in items:
                raise ValueError(f"{location}: пустой или повторяющийся артикул «{sku}»")
            try:
                items[sku] = parse(row[value_column])
            except ValueError as error:
                raise ValueError(f"{location}: {error}") from error
    if not items:
        raise ValueError(f"{path}: нет товарных позиций")
    return items


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


def main():
    parser = argparse.ArgumentParser(description="Расчёт заявки по прайсу поставщика")
    parser.add_argument("request", type=Path, help="Заявка (.csv или .xlsx)")
    parser.add_argument("prices", type=Path, help="Прайс (.csv или .xlsx)")
    args = parser.parse_args()
    try:
        request = read_items(args.request, "количество", quantity)
        prices = read_items(args.prices, "цена", price)
    except (ValueError, OSError, csv.Error) as error:
        parser.exit(2, f"Ошибка: {error}\n")
    print(report(request, prices))


if __name__ == "__main__":
    main()
