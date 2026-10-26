@echo off
rem Запуск на Windows двойным щелчком: при первом запуске ставит зависимости
cd /d %~dp0
if not exist .venv (
  py -3 -m venv .venv
  .venv\Scripts\python -m pip install -r requirements.txt
)
.venv\Scripts\python -m streamlit run app.py
