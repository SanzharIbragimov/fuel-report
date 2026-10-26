"""Веб-интерфейс: streamlit run app.py"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from fuel.core import COLUMNS, Settings, analyze, detect_columns, read_any, read_rates, summary
from fuel.report import build_excel

st.set_page_config(page_title="Расход топлива — отчёт", page_icon="⛽", layout="wide")
LABELS = {"plate": "Госномер", "model": "Модель", "driver": "Водитель", "date": "Дата", "mileage": "Пробег, км",
          "fact": "Факт, л", "norm_l": "Норма, л", "norm_rate": "Норма, л/100 км"}
SAMPLE = Path(__file__).parent / "samples" / "primer_avtopark_oktyabr.xlsx"

st.title("⛽ Отчёт по расходу топлива")
st.caption("Загрузите выгрузку телематики или путевых листов (.xlsx / .csv). Инструмент сопоставит факт с нормой, "
           "посчитает отклонения и соберёт отчёт в Excel.")

with st.sidebar:
    st.header("Настройки")
    threshold_pct = st.number_input("Порог перерасхода, %", 0.0, 100.0, 10.0, 1.0)
    threshold_l = st.number_input("Не учитывать превышение меньше, л", 0.0, 100.0, 5.0, 1.0)
    winter = st.number_input("Зимняя надбавка к норме, % (ноябрь–март)", 0.0, 50.0, 0.0, 1.0,
                             help="Если надбавка установлена приказом предприятия. 0 — не применять.")
    fix_typos = st.checkbox("Исправлять опечатки в госномерах", True)
    rates_file = st.file_uploader("Справочник норм (необязательно)", type=["xlsx", "csv"],
                                  help="Колонки «Модель» или «Госномер» и «Норма, л/100 км» — если в основном файле нормы нет")

up = st.file_uploader("Файл с данными автопарка", type=["xlsx", "xls", "csv"])
use_sample = st.button("Попробовать на примере") if not up else False
if use_sample:
    st.session_state["sample"] = True
if not up and not st.session_state.get("sample"):
    st.info("Нужны колонки: госномер, пробег, фактический расход и норма (литры или л/100 км). "
            "Водитель, модель и дата — по желанию. Названия колонок распознаются автоматически.")
    st.stop()

data, name = (up.getvalue(), up.name) if up else (SAMPLE.read_bytes(), SAMPLE.name)
sheet = 0
if not name.lower().endswith(".csv"):
    sheets = pd.ExcelFile(pd.io.common.BytesIO(data)).sheet_names
    if len(sheets) > 1:
        sheet = st.selectbox("Лист", sheets)
df = read_any(data, name, sheet)
st.caption(f"Файл: **{name}** · строк: {len(df)}")

# --- соответствие колонок: автоматически, с возможностью поправить ---
auto = detect_columns(df)
with st.expander("Соответствие колонок", expanded=not all(k in auto for k in ("plate", "mileage", "fact"))):
    cols = st.columns(4)
    mapping = {}
    options = ["—"] + list(df.columns)
    for i, key in enumerate(LABELS):
        default = auto.get(key)
        choice = cols[i % 4].selectbox(LABELS[key], options, index=options.index(default) if default in options else 0,
                                       key=f"map_{key}")
        if choice != "—":
            mapping[key] = choice

settings = Settings(threshold_pct=threshold_pct, threshold_l=threshold_l, winter_coef_pct=winter, fix_plate_typos=fix_typos)
if rates_file:
    try:
        settings.rates = read_rates(read_any(rates_file.getvalue(), rates_file.name))
        st.sidebar.success(f"Норм в справочнике: {len(settings.rates)}")
    except ValueError as e:
        st.sidebar.error(str(e))

try:
    det, issues, tot = analyze(df, settings, mapping)
except ValueError as e:
    st.error(str(e))
    st.stop()

# --- итоги ---
def num(x):
    return f"{x:,.0f}".replace(",", " ")


c = st.columns(6)
c[0].metric("Машин", tot["vehicles"])
c[1].metric("Факт, л", num(tot["fact"]))
c[2].metric("Норма, л", num(tot["norm"]))
c[3].metric("Отклонение", f"{tot['dev_pct']:+.1f}%" if tot["dev_pct"] is not None else "—")
c[4].metric("Перерасход, л", num(tot["overrun_l"]))
c[5].metric("Проблемных записей", tot["problems"])

if issues:
    errs = [x for x in issues if x.level == "ошибка"]
    st.warning(f"Во входных данных: ошибок — {len(errs)} (строки не учтены), предупреждений — {len(issues) - len(errs)}. "
               "Подробности — на вкладке «Ошибки в данных».")

xlsx = build_excel(det, issues, tot, settings, name)
st.download_button("⬇ Скачать отчёт в Excel", xlsx, file_name=f"otchet_toplivo_{Path(name).stem}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

tabs = st.tabs(["Проблемы", "По машинам", "По водителям", "Детально", "Ошибки в данных"])
show = ["Строка", "Дата", "Госномер", "Модель", "Водитель", "Пробег, км", "Факт, л", "Норма, л",
        "Отклонение, л", "Отклонение, %", "Признак"]
fmt = {"Дата": st.column_config.DateColumn(format="DD.MM.YYYY"),
       "Отклонение, %": st.column_config.NumberColumn(format="%+.1f%%")}


def paint(row):
    v = row.get("Отклонение, %")
    if pd.isna(v):
        return [""] * len(row)
    color = "#f8d7da" if v > threshold_pct else ("#fff3cd" if v > 0 else "")
    return [f"background-color: {color}" if color else ""] * len(row)


with tabs[0]:
    prob = det[det["Проблема"]].sort_values("Отклонение, %", ascending=False)[show]
    st.caption("Перерасход выше порога, расход без пробега и другие записи, которые стоит проверить.")
    st.dataframe(prob, hide_index=True, width="stretch", column_config=fmt)
with tabs[1]:
    s = summary(det, "Госномер")
    st.dataframe(s.style.apply(paint, axis=1).format(precision=1), hide_index=True, width="stretch", column_config=fmt)
    if len(s):
        st.bar_chart(s.set_index("Госномер")["Отклонение, %"], y_label="Отклонение от нормы, %", horizontal=True)
with tabs[2]:
    s = summary(det, "Водитель")
    if len(s):
        st.dataframe(s.style.apply(paint, axis=1).format(precision=1), hide_index=True, width="stretch", column_config=fmt)
    else:
        st.info("В файле нет колонки с водителем.")
with tabs[3]:
    st.dataframe(det[show].style.apply(paint, axis=1).format(precision=1), hide_index=True, width="stretch", column_config=fmt)
with tabs[4]:
    if issues:
        st.dataframe(pd.DataFrame([(x.row, x.level, x.column, x.message) for x in sorted(issues, key=lambda x: x.row)],
                                  columns=["Строка файла", "Уровень", "Поле", "Что не так"]),
                     hide_index=True, width="stretch")
    else:
        st.success("Ошибок во входных данных нет.")
