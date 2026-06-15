@echo off
cd /d "%~dp0"
title White Line Trucking Management Portal - Office Network Share

echo Starting White Line Trucking Management Portal for office network sharing...
echo.
echo The main computer can open: http://localhost:8501
echo Other computers on the same Wi-Fi/network must use the Network URL shown below.
echo It will usually look like: http://192.168.x.x:8501
echo.
echo Leave this window open while anyone is using the tool.
echo.

start "" cmd /c "timeout /t 7 >nul & start "" http://localhost:8501"

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501 --server.headless false

pause
