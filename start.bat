@echo off
REM 财报跟踪平台可移植版 · Windows 一键启动
cd /d "%~dp0"
pip install -r requirements.txt --quiet 2>nul
python run.py %*
