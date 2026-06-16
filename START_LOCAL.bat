@echo off
cd /d "%~dp0"
title White Line Trucking Management Portal

echo Starting White Line Trucking Management Portal...
echo.
echo This will open in your browser at: http://localhost:8501
echo Leave this window open while using the tool.
echo.

REM Open browser after Streamlit has a few seconds to start.
start "" cmd /c "timeout /t 7 >nul & start "" http://localhost:8501"

REM Install required packages and start the app.
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py --server.address localhost --server.port 8501 --server.headless false

echo.
echo If the browser did not open, manually open: http://localhost:8501
pause
