"""Генерирует тестовый файл автопарка с типичными ошибками и аномалиями.

    python scripts/make_sample.py  →  samples/primer_avtopark_oktyabr.xlsx, samples/normy.xlsx
Все данные вымышленные.
"""
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

OUT = Path(__file__).resolve().parent.parent / "samples"
FLEET = [  # госномер, модель, норма л/100 км, водитель, типичный пробег за смену
    ("101ABC01", "ГАЗель Next", 16.0, "Ахметов Ержан", 180),
    ("214KTE01", "ГАЗель Next", 16.0, "Сейткали Нурлан", 160),
    ("305MNA01", "КамАЗ 65115", 34.0, "Иванов Сергей", 220),
    ("417PRT01", "КамАЗ 65115", 34.0, "Жумабаев Асхат", 240),
    ("522HBX02", "Toyota Hilux", 12.5, "Ким Виталий", 140),
    ("631ATP02", "Lada Largus", 9.0, "Омаров Дамир", 120),
    ("748YKM01", "Shacman X3000", 38.0, "Бекжанов Талгат", 260),
    ("856CEA01", "Hyundai HD78", 18.5, "Нуртаев Болат", 170),
    ("963OXT01", "Toyota Hilux", 12.5, "Абенов Даурен", 150),
    ("079MKP01", "ГАЗель Next", 16.0, "Касымов Арман", 175),
]


def main():
    rnd = random.Random(42)
    OUT.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Октябрь"
    ws.append(["Гос. номер", "Марка/модель", "Водитель", "Дата", "Пробег, км", "Расход факт, л", "Норма, л/100 км"])
    start = date(2026, 10, 1)
    rows = []
    for d in range(26):
        day = start + timedelta(days=d)
        if day.weekday() == 6:
            continue
        for plate, model, rate, driver, km in FLEET:
            if rnd.random() < 0.12:
                continue
            mileage = max(10, round(rnd.gauss(km, km * 0.25)))
            fact = mileage * rate / 100 * rnd.uniform(0.93, 1.07)
            rows.append([plate, model, driver, day.strftime("%d.%m.%Y"), mileage, round(fact, 1), rate])
    # аномалии: систематический перерасход одного водителя, разовые всплески, «слив»
    for r in rows:
        if r[0] == "305MNA01" and rnd.random() < 0.7:
            r[5] = round(r[5] * rnd.uniform(1.18, 1.32), 1)
        if r[0] in ("748YKM01", "856CEA01") and rnd.random() < 0.08:
            r[5] = round(r[5] * 1.45, 1)
    rows[40][4], rows[40][5] = 0, 42.0                             # расход без пробега
    rows[95][4], rows[95][5] = 0, 55.0
    # типичные ошибки ввода
    rows[12][0] = "101 АВС 01"                                     # пробелы и кириллица — исправится молча
    rows[33][0] = "214KTF01"                                       # опечатка в номере
    rows[51][3] = "31.09.2026"                                     # несуществующая дата
    rows[64][4] = ""                                               # пустой пробег
    rows[77][5] = "сорок"                                          # текст вместо числа
    rows[88][3] = "2026-10-14"                                     # другой формат даты — распознается
    rows[101][5] = "38,5"                                          # запятая — распознается
    rows[118][0] = ""                                              # пустой номер
    for r in rows:
        ws.append(r)
    for col, w in zip("ABCDEFG", (14, 18, 22, 12, 12, 16, 16)):
        ws.column_dimensions[col].width = w
    wb.save(OUT / "primer_avtopark_oktyabr.xlsx")

    # справочник норм — для файлов без колонки нормы
    wb = Workbook()
    ws = wb.active
    ws.title = "Нормы"
    ws.append(["Модель", "Норма, л/100 км"])
    for m, r in sorted({(f[1], f[2]) for f in FLEET}):
        ws.append([m, r])
    wb.save(OUT / "normy.xlsx")
    print(f"{len(rows)} строк → {OUT}")


if __name__ == "__main__":
    main()
