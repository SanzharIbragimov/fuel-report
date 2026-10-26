"""Без интерфейса: python cli.py данные.xlsx [--threshold 10] [--norms normy.xlsx] [--out отчёт.xlsx]"""
import argparse
from pathlib import Path

from fuel.core import Settings, analyze, read_any, read_rates
from fuel.report import build_excel

p = argparse.ArgumentParser()
p.add_argument("file")
p.add_argument("--threshold", type=float, default=10.0, help="порог перерасхода, %")
p.add_argument("--min-liters", type=float, default=5.0)
p.add_argument("--winter", type=float, default=0.0, help="зимняя надбавка к норме, %")
p.add_argument("--norms", help="справочник норм .xlsx/.csv")
p.add_argument("--out")
a = p.parse_args()
src = Path(a.file)
st = Settings(threshold_pct=a.threshold, threshold_l=a.min_liters, winter_coef_pct=a.winter)
if a.norms:
    st.rates = read_rates(read_any(Path(a.norms).read_bytes(), a.norms))
det, issues, tot = analyze(read_any(src.read_bytes(), src.name), st)
out = Path(a.out or f"otchet_{src.stem}.xlsx")
out.write_bytes(build_excel(det, issues, tot, st, src.name))
print(f"Машин: {tot['vehicles']}, записей: {tot['rows_ok']}/{tot['rows_in']}, проблемных: {tot['problems']}, "
      f"ошибок в данных: {tot['errors']} → {out}")
