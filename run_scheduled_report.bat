@echo off
cd /d "%~dp0"
"C:\Users\lenovo\AppData\Local\Programs\Python\Python312\python.exe" run_scheduled_report.py >> logs\scheduled_run.log 2>&1
