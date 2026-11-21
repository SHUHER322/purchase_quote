import sys
import tempfile
import unittest
from decimal import Decimal
from itertools import product
from pathlib import Path
from random import Random
from subprocess import run

from allocation import Offer, allocate
from purchase import read_offers


class AllocationTests(unittest.TestCase):
    def test_small_request_prefers_smaller_pack(self):
        bulk = Offer("Коробки", Decimal("2.01"), 10)
        retail = Offer("Поштучно", Decimal("3.99"))
        orders = allocate(3, [bulk, retail])
        self.assertEqual([(o.offer, o.units, o.amount) for o in orders],
                         [(retail, 3, Decimal("11.97"))])

    def test_split_order_rounds_stock_down_to_full_packs(self):
        offers = [Offer("Первый", Decimal("1"), 3, 8), Offer("Второй", Decimal("2"), 4, 8)]
        orders = allocate(9, offers)
        self.assertEqual([(o.offer.supplier, o.units) for o in orders],
                         [("Первый", 6), ("Второй", 4)])
        self.assertEqual(sum(o.amount for o in orders), 14)

    def test_matches_exhaustive_search(self):
        random = Random(731)
        for case in range(200):
            demand = random.randint(1, 20)
            offers = [Offer(str(i), Decimal(random.randrange(5)), random.randint(1, 6),
                            random.randrange(15)) for i in range(3)]
            scores = []
            for counts in product(*(range(o.stock // o.pack + 1) for o in offers)):
                units = sum(n * o.pack for n, o in zip(counts, offers))
                cost = sum(n * o.pack * o.price for n, o in zip(counts, offers))
                scores.append((max(0, demand - units), cost, units))
            orders = allocate(demand, offers)
            with self.subTest(case=case):
                for order in orders:
                    self.assertLessEqual(order.units, order.offer.stock)
                units = sum(o.units for o in orders)
                self.assertEqual((max(0, demand - units), sum(o.amount for o in orders), units),
                                 min(scores))

    def test_invalid_stock_and_pack_report_row(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "Прайс.csv"
            for stock, pack in [("", "1"), ("-1", "2"), ("2", "0"), ("2", "1.5")]:
                path.write_text(f"артикул;цена;остаток, шт;штук в упаковке\nA;5;{stock};{pack}\n",
                                encoding="utf-8")
                with self.subTest(stock=stock, pack=pack), self.assertRaises(ValueError) as error:
                    read_offers(path)
                self.assertIn(f"{path}, строка 2", str(error.exception))

    def test_multi_supplier_cli(self):
        root = Path(__file__).resolve().parents[1]
        result = run([sys.executable, "-X", "utf8", "purchase.py", "examples/request.csv",
                      "examples/prices.xlsx", "examples/Склад.csv"],
                     cwd=root, capture_output=True, encoding="utf-8", check=True)
        self.assertIn("Склад: 2 уп. × 2 шт.", result.stdout)
        self.assertIn("LABEL-03: заявка 2, заказать 1, излишек 0, дефицит 1", result.stdout)
        self.assertTrue(result.stdout.rstrip().endswith("Неполный заказ (позиций с дефицитом: 1): 652,10 ₽"))
