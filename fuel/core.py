"""Загрузка, проверка и расчёт отклонений расхода топлива от норм.

Вход — Excel/CSV с поездками или сменами: госномер, водитель, дата, пробег, факт и норма.
Норма может быть задана в файле литрами, в л/100 км или взята из справочника норм по модели.
"""
from __future__ import annotations

import difflib
import io
import re
from collections import Counter
from dataclasses import dataclass, field

import pandas as pd

# --- распознавание колонок по названию (рус., каз., англ.) ---
COLUMNS: dict[str, list[str]] = {
    "plate": ["госномер", "гос номер", "гос. номер", "гос.номер", "номер тс", "номер авто", "мемлекеттік нөмір",
              "нөмір", "plate", "vehicle", "тс", "автомобиль", "транспорт"],
    "model": ["модель", "марка", "марка/модель", "марка тс", "model", "үлгі"],
    "driver": ["водитель", "фио водителя", "жүргізуші", "driver"],
    "date": ["дата", "дата поездки", "период", "күні", "date", "day"],
    "mileage": ["пробег", "пробег, км", "пробег км", "км", "жүріс", "mileage", "distance", "odometer delta"],
    "fact": ["факт", "фактический расход", "расход факт", "факт, л", "расход, л", "израсходовано",
             "нақты шығын", "fuel used", "actual"],
    "norm_l": ["норма, л", "норма л", "нормативный расход", "расход по норме", "норма расхода, л",
               "норма бойынша", "norm liters", "norm"],
    "norm_rate": ["норма л/100", "норма, л/100 км", "линейная норма", "л/100 км", "норма на 100 км",
                  "норма л/100км", "rate", "l/100km"],
    "refuel": ["заправка", "заправлено", "заправка, л", "құйылды", "refuel"],
}
REQUIRED = ("plate", "mileage", "fact")

# кириллица, похожая на латиницу, — в госномерах РК используется латиница
_LOOKALIKE = str.maketrans("АВЕКМНОРСТХУ", "ABEKMHOPCTXY")


@dataclass
class Settings:
    threshold_pct: float = 10.0          # перерасход выше этого — проблема
    threshold_l: float = 5.0             # и не меньше этого в литрах (отсекает шум на коротких поездках)
    winter_coef_pct: float = 0.0         # зимняя надбавка к норме, %, если задана приказом предприятия
    winter_months: tuple = (11, 12, 1, 2, 3)
    fix_plate_typos: bool = True         # объединять очевидные опечатки госномеров
    rates: dict = field(default_factory=dict)   # справочник: модель или госномер → л/100 км


@dataclass
class Issue:
    row: int          # номер строки в исходном файле (как в Excel)
    level: str        # ошибка | предупреждение
    column: str
    message: str


# ---------------- загрузка ----------------
def read_any(data: bytes, name: str, sheet: str | int | None = 0) -> pd.DataFrame:
    if name.lower().endswith(".csv"):
        text = data.decode("utf-8-sig", errors="replace")
        sep = ";" if text.splitlines()[0].count(";") > text.splitlines()[0].count(",") else ","
        return pd.read_csv(io.StringIO(text), sep=sep, dtype=str, keep_default_na=False)
    return pd.read_excel(io.BytesIO(data), sheet_name=sheet, dtype=object)


def _norm_header(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower().replace("ё", "е"))


def detect_columns(df: pd.DataFrame) -> dict[str, str]:
    """Сопоставляет колонки файла с полями. Сначала точное совпадение, потом по вхождению."""
    found: dict[str, str] = {}
    headers = {c: _norm_header(c) for c in df.columns}
    for key, names in COLUMNS.items():
        for c, h in headers.items():
            if c not in found.values() and h in names:
                found[key] = c
                break
    for key, names in COLUMNS.items():
        if key in found:
            continue
        for c, h in headers.items():
            if c in found.values():
                continue
            if any(n in h for n in names if len(n) > 3):
                # «норма, л/100 км» не должна попасть в «норма, л»
                if key == "norm_l" and ("100" in h):
                    continue
                found[key] = c
                break
    return found


# ---------------- очистка значений ----------------
def blank(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and v.strip() == "") \
        or v is pd.NaT


def text(v) -> str:
    return "" if blank(v) else str(v).strip()


def normalize_plate(v) -> str:
    s = text(v).upper().translate(_LOOKALIKE)
    return re.sub(r"[^A-Z0-9]", "", s)


def to_number(v) -> float | None:
    if blank(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(" ", "").replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    if s in ("", "-", ".") or s.count(".") > 1:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def to_date(v):
    if blank(v):
        return None
    if isinstance(v, (pd.Timestamp,)) or hasattr(v, "year"):
        return pd.Timestamp(v).normalize()
    if isinstance(v, (int, float)) and 30000 < v < 60000:          # дата Excel числом
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(v))
    s = str(v).strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return pd.Timestamp(pd.to_datetime(s, format=fmt)).normalize()
        except (ValueError, TypeError):
            continue
    return None


def plate_typos(plates: list[str]) -> dict[str, str]:
    """Редкий номер, отличающийся от частого одним символом, — вероятная опечатка."""
    counts = Counter(p for p in plates if p)
    fixes = {}
    frequent = [p for p, n in counts.most_common() if n >= 2]
    for p, n in counts.items():
        if n > 1 or p in frequent:
            continue
        match = difflib.get_close_matches(p, frequent, n=1, cutoff=0.85)
        if match and len(match[0]) == len(p) and sum(a != b for a, b in zip(match[0], p)) == 1:
            fixes[p] = match[0]
    return fixes


# ---------------- основной расчёт ----------------
def analyze(df: pd.DataFrame, settings: Settings | None = None, mapping: dict[str, str] | None = None):
    """Возвращает (детальная таблица, список проблем во входных данных, итоговые цифры)."""
    st = settings or Settings()
    mapping = mapping or detect_columns(df)
    issues: list[Issue] = []
    missing = [k for k in REQUIRED if k not in mapping]
    if missing:
        names = {"plate": "госномер", "mileage": "пробег", "fact": "фактический расход"}
        raise ValueError("Не найдены колонки: " + ", ".join(names[m] for m in missing)
                         + ". Переименуйте их или укажите соответствие вручную.")
    if "norm_l" not in mapping and "norm_rate" not in mapping and not st.rates:
        raise ValueError("Нет нормы расхода: нужна колонка «Норма, л» или «Норма, л/100 км», либо справочник норм.")

    rows = []
    for i, r in df.iterrows():
        excel_row = int(i) + 2                     # строка 1 — заголовок
        raw_plate = r.get(mapping["plate"])
        rec = {
            "Строка": excel_row,
            "Госномер": normalize_plate(raw_plate),
            "Модель": text(r.get(mapping["model"])) if "model" in mapping else "",
            "Водитель": re.sub(r"\s+", " ", text(r.get(mapping["driver"]))) if "driver" in mapping else "",
            "Дата": to_date(r.get(mapping["date"])) if "date" in mapping else None,
            "Пробег, км": to_number(r.get(mapping["mileage"])),
            "Факт, л": to_number(r.get(mapping["fact"])),
            "Норма, л": to_number(r.get(mapping["norm_l"])) if "norm_l" in mapping else None,
            "Норма, л/100 км": to_number(r.get(mapping["norm_rate"])) if "norm_rate" in mapping else None,
        }
        if all(rec[k] in (None, "") for k in ("Госномер", "Пробег, км", "Факт, л")):
            continue                                # пустая строка — пропускаем молча
        # --- проверки ---
        if not rec["Госномер"]:
            issues.append(Issue(excel_row, "ошибка", "Госномер", "пустой госномер — строка не учтена"))
            continue
        if "date" in mapping and rec["Дата"] is None and not blank(r.get(mapping["date"])):
            issues.append(Issue(excel_row, "предупреждение", "Дата", f"неверный формат даты «{r.get(mapping['date'])}»"))
        elif "date" in mapping and rec["Дата"] is None:
            issues.append(Issue(excel_row, "предупреждение", "Дата", "пустая дата"))
        for col in ("Пробег, км", "Факт, л"):
            if rec[col] is None:
                src = r.get(mapping["mileage" if col.startswith("Пробег") else "fact"])
                msg = "пустая ячейка" if blank(src) else f"не число «{src}»"
                issues.append(Issue(excel_row, "ошибка", col, msg + " — строка не учтена"))
            elif rec[col] < 0:
                issues.append(Issue(excel_row, "ошибка", col, f"отрицательное значение {rec[col]:g} — строка не учтена"))
                rec[col] = None
        if rec["Пробег, км"] is None or rec["Факт, л"] is None:
            continue
        # --- норма ---
        rate = rec["Норма, л/100 км"] or st.rates.get(rec["Госномер"]) or st.rates.get(rec["Модель"])
        norm = rec["Норма, л"]
        if norm is None and rate is not None:
            norm = rec["Пробег, км"] * rate / 100
        if norm is None:
            issues.append(Issue(excel_row, "ошибка", "Норма", "нет нормы для машины — строка не учтена"))
            continue
        if st.winter_coef_pct and rec["Дата"] is not None and rec["Дата"].month in st.winter_months:
            norm *= 1 + st.winter_coef_pct / 100
        rec["Норма, л"] = round(norm, 2)
        rec["Норма, л/100 км"] = rate
        rows.append(rec)

    det = pd.DataFrame(rows, columns=["Строка", "Госномер", "Модель", "Водитель", "Дата", "Пробег, км",
                                      "Факт, л", "Норма, л/100 км", "Норма, л"])
    # --- опечатки в госномерах ---
    if st.fix_plate_typos and len(det):
        fixes = plate_typos(det["Госномер"].tolist())
        for wrong, right in fixes.items():
            for rownum in det.loc[det["Госномер"] == wrong, "Строка"]:
                issues.append(Issue(int(rownum), "предупреждение", "Госномер",
                                    f"вероятная опечатка «{wrong}» — учтено как {right}"))
        det["Госномер"] = det["Госномер"].replace(fixes)

    det["Отклонение, л"] = (det["Факт, л"] - det["Норма, л"]).round(2)
    det["Отклонение, %"] = ((det["Факт, л"] / det["Норма, л"] - 1) * 100).where(det["Норма, л"] > 0).round(1)
    det["Признак"] = [flag(r, st) for r in det.itertuples(index=False)]
    det["Проблема"] = det["Признак"] != ""
    totals = {
        "rows_in": int(len(df)), "rows_ok": int(len(det)),
        "errors": sum(1 for x in issues if x.level == "ошибка"),
        "warnings": sum(1 for x in issues if x.level == "предупреждение"),
        "vehicles": int(det["Госномер"].nunique()) if len(det) else 0,
        "fact": float(det["Факт, л"].sum()), "norm": float(det["Норма, л"].sum()),
        "overrun_l": float(det.loc[det["Отклонение, л"] > 0, "Отклонение, л"].sum()),
        "problems": int(det["Проблема"].sum()),
    }
    totals["dev_pct"] = (totals["fact"] / totals["norm"] - 1) * 100 if totals["norm"] else None
    return det, issues, totals


def flag(r, st: Settings) -> str:
    """Причина, по которой запись попадает в список проблем."""
    mileage, fact, dev_l, dev_pct = r[5], r[6], r[9], r[10]
    if mileage == 0 and fact > 0:
        return "расход без пробега — проверить на слив / нецелевое использование"
    if mileage > 0 and fact == 0:
        return "пробег без расхода — проверить данные"
    if dev_pct is not None and dev_pct > st.threshold_pct and dev_l >= st.threshold_l:
        return f"перерасход {dev_pct:.0f}% (+{dev_l:.0f} л)"
    if dev_pct is not None and dev_pct < -40:
        return f"расход на {abs(dev_pct):.0f}% ниже нормы — проверить пробег/показания"
    return ""


def summary(det: pd.DataFrame, by: str) -> pd.DataFrame:
    """Сводка по машинам или водителям, проблемные сверху."""
    if not len(det) or by not in det:
        return pd.DataFrame()
    g = det.groupby(det[by].replace("", "не указан"), dropna=False).agg(
        **{"Записей": ("Строка", "count"), "Пробег, км": ("Пробег, км", "sum"),
           "Факт, л": ("Факт, л", "sum"), "Норма, л": ("Норма, л", "sum"),
           "Проблемных записей": ("Проблема", "sum")}).reset_index()
    if by == "Госномер" and "Модель" in det:
        models = det.groupby("Госномер")["Модель"].agg(lambda s: next((x for x in s if x), ""))
        g.insert(1, "Модель", g["Госномер"].map(models).fillna(""))
    g["Отклонение, л"] = (g["Факт, л"] - g["Норма, л"]).round(2)
    g["Отклонение, %"] = ((g["Факт, л"] / g["Норма, л"] - 1) * 100).where(g["Норма, л"] > 0).round(1)
    g["Факт, л/100 км"] = (g["Факт, л"] / g["Пробег, км"] * 100).where(g["Пробег, км"] > 0).round(1)
    for c in ("Пробег, км", "Факт, л", "Норма, л"):
        g[c] = g[c].round(1)
    return g.sort_values(["Проблемных записей", "Отклонение, %"], ascending=False).reset_index(drop=True)


def read_rates(df: pd.DataFrame) -> dict:
    """Справочник норм: колонки «Модель» или «Госномер» и «Норма, л/100 км»."""
    m = detect_columns(df)
    key = m.get("plate") or m.get("model")
    val = m.get("norm_rate") or m.get("norm_l")
    if not key or not val:
        raise ValueError("В справочнике нужны колонки «Модель» (или «Госномер») и «Норма, л/100 км».")
    out = {}
    for _, r in df.iterrows():
        k, v = r.get(key), to_number(r.get(val))
        if k is not None and v:
            out[normalize_plate(k) if key == m.get("plate") else str(k).strip()] = v
    return out
