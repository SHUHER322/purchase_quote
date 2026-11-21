import argparse
import csv
import re
from decimal import Decimal
from pathlib import Path


def read_items(path, field, parse):
    items = {}
    with path.open(encoding="utf-8-sig", newline="") as source:
        reader = csv.reader(source, delimiter=";", strict=True)
        headers = next(reader, [])
        if headers.count("артикул") != 1 or headers.count(field) != 1:
            raise ValueError(f"{path}: нужны столбцы «артикул» и «{field}»")
        for row in reader:
            if not row:
                continue
            location = f"{path}, строка {reader.line_num}"
            if len(row) != len(headers):
                raise ValueError(f"{location}: проверьте число полей")
            sku = row[headers.index("артикул")].strip()
            if not sku or sku in items:
                raise ValueError(f"{location}: пустой или повторяющийся артикул «{sku}»")
            try:
                items[sku] = parse(row[headers.index(field)].strip())
            except ValueError as error:
                raise ValueError(f"{location}: {error}") from error
    if not items:
        raise ValueError(f"{path}: нет товарных позиций")
    return items


def quantity(value):
    if not re.fullmatch(r"[0-9]+", value) or int(value) < 1:
        raise ValueError("количество должно быть целым положительным числом")
    return int(value)


def price(value):
    if not re.fullmatch(r"[0-9]+([.,][0-9]{1,2})?", value):
        raise ValueError("нужна неотрицательная цена с точностью до копеек")
    return Decimal(value.replace(",", "."))


def rubles(value):
    return f"{value:.2f}".replace(".", ",") + " ₽"


def report(request, prices):
    lines = []
    total = Decimal(0)
    missing = 0
    for sku, count in request.items():
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
    parser.add_argument("request", type=Path, help="CSV с заявкой")
    parser.add_argument("prices", type=Path, help="CSV с ценами")
    args = parser.parse_args()
    try:
        request = read_items(args.request, "количество", quantity)
        prices = read_items(args.prices, "цена", price)
    except (ValueError, OSError, csv.Error) as error:
        parser.exit(2, f"Ошибка: {error}\n")
    print(report(request, prices))


if __name__ == "__main__":
    main()
