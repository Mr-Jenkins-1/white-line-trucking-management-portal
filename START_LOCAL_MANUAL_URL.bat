@echo off
cd /d "%~dp0"
title White Line Trucking Management Portal - Manual URL

echo Starting White Line Trucking Management Portal...
echo.
echo If your browser does not open automatically, copy and paste this into Chrome or Edge:
echo.
echo http://localhost:8501
echo.
echo Leave this window open while using the tool.
echo.

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py --server.address localhost --server.port 8501 --server.headless true

pause
