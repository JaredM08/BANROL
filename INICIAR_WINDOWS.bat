@echo off
cd /d "%~dp0"
echo Instalando (solo la primera vez)...
pip install -r requirements.txt >nul 2>&1
echo Abriendo el programa en tu navegador...
start "" http://localhost:8000
python banrol_motor.py --serve --port 8000
pause
