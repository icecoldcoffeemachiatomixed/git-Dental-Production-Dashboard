@echo off
cd /d "%~dp0"
set DENTAL_DASHBOARD_LOCAL=1
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" -m streamlit run app.py
