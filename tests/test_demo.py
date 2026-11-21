import sys
import tempfile
import unittest
from contextlib import closing
from decimal import Decimal
from pathlib import Path
from subprocess import run

from openpyxl import load_workbook


class DemoTests(unittest.TestCase):
    def test_three_supplier_workbook(self):
        root = Path(__file__).resolve().parents[1]
        demo = root / "examples" / "demo"
        # Контрольные значения рассчитаны вручную по трём демонстрационным прайсам.
        # Сводка: запрошено, заказано, излишек, дефицит и сумма в рублях.
        summary = [
            ("BOX-01", 4, 4, 0, 0, 322.20), ("TAPE-02", 3, 3, 0, 0, 180),
            ("BAG-03", 10, 10, 0, 0, 44), ("MARK-04", 2, 2, 0, 0, 36),
            ("FILM-05", 3, 4, 1, 0, 84), ("LABEL-06", 5, 2, 0, 3, 22),
        ]
        # Заказы: артикул, упаковки, фасовка, штуки, цена и сумма.
        orders = {
            "Картонка": [("BAG-03", 3, 2, 6, 4, 24), ("FILM-05", 1, 4, 4, 21, 84)],
            "Пакетика": [("BOX-01", 2, 2, 4, 80.55, 322.20), ("BAG-03", 2, 2, 4, 5, 20),
                         ("LABEL-06", 1, 2, 2, 11, 22)],
            "Рулоника": [("TAPE-02", 3, 1, 3, 60, 180), ("MARK-04", 2, 1, 2, 18, 36)],
        }
        inputs = [demo / f"{name}.csv" for name in ("Заявка", *orders)]
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "демо-заказ.xlsx"
            run([sys.executable, "-X", "utf8", root / "purchase.py", *inputs, "-o", output],
                check=True, capture_output=True)
            for path in (output, demo / "Отчёт.xlsx"):
                with self.subTest(path=path.name), closing(load_workbook(path, data_only=True)) as book:
                    self.assertEqual(set(book.sheetnames), {"Сводка", "Требуют внимания", *orders})
                    sheet = book["Сводка"]
                    self.assertEqual(list(sheet.iter_rows(min_row=4, max_col=6, values_only=True)), summary)
                    self.assertIn("дефицитом: 1", sheet["C2"].value)
                    attention = book["Требуют внимания"]
                    self.assertEqual(list(attention.iter_rows(min_row=4, max_col=6, values_only=True)),
                                     summary[-2:])
                    total = Decimal(0)
                    for supplier, expected in orders.items():
                        sheet = book[supplier]
                        actual = [(row[0], *row[2:]) for row in sheet.iter_rows(min_row=4, values_only=True)]
                        self.assertEqual(actual, expected)
                        subtotal = sum(Decimal(str(row[-1])) for row in actual)
                        self.assertEqual(subtotal, Decimal(str(sheet["B2"].value)))
                        total += subtotal
                    self.assertEqual(total, Decimal("688.20"))
                    self.assertEqual(total, Decimal(str(book["Сводка"]["B2"].value)))
