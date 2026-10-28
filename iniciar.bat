@echo off
REM Savian - Sistema de Fechamento DRE (Windows)
cd /d %~dp0
if not exist .venv (
  py -m venv .venv
  .venv\Scripts\pip install -r requirements.txt
)
.venv\Scripts\streamlit run app.py
