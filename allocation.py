"""Подбор целых упаковок для одного артикула."""
from dataclasses import dataclass
from decimal import Decimal


@dataclass(frozen=True)
class Offer:
    supplier: str
    price: Decimal
    pack: int = 1
    stock: int | None = None

    def available_packs(self, demand):
        needed = (demand + self.pack - 1) // self.pack
        return needed if self.stock is None else min(needed, self.stock // self.pack)


@dataclass(frozen=True)
class Order:
    offer: Offer
    packs: int

    @property
    def units(self):
        return self.packs * self.offer.pack

    @property
    def amount(self):
        return self.units * self.offer.price


def allocate(demand: int, offers: list[Offer]) -> list[Order]:
    available = [(offer, offer.available_packs(demand)) for offer in offers]
    available = [(offer, count) for offer, count in available if count]
    if sum(offer.pack * count for offer, count in available) <= demand:
        return [Order(offer, count) for offer, count in available]

    # При одинаковой фасовке достаточно брать доступные упаковки по возрастанию цены.
    if len({offer.pack for offer, _ in available}) == 1:
        orders = []
        for offer, count in sorted(available, key=lambda item: item[0].price):
            order = Order(offer, min(count, (demand + offer.pack - 1) // offer.pack))
            orders.append(order)
            demand -= order.units
            if demand <= 0:
                break
        return orders

    # Для каждого покрытия храним минимальную сумму, затем минимальное число
    # купленных штук. Все варианты с полным покрытием сравниваются между собой.
    states = {0: (Decimal(0), 0, None)}
    for offer, count in available:
        # Группы 1, 2, 4, ... позволяют набрать любое допустимое число упаковок.
        group = 1
        while count:
            packs = min(group, count)
            count -= packs
            group *= 2
            order = Order(offer, packs)
            for covered, (cost, units, previous) in list(states.items()):
                if covered == demand:
                    continue
                new_units = units + order.units
                key = min(demand, new_units)
                score = (cost + order.amount, new_units)
                if key not in states or score < states[key][:2]:
                    states[key] = (*score, (order, previous))

    counts = {}
    node = states[demand][2]
    while node is not None:
        order, node = node
        counts[order.offer] = counts.get(order.offer, 0) + order.packs
    return [Order(offer, counts[offer]) for offer, _ in available if offer in counts]
