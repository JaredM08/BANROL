#!/bin/bash
cd "$(dirname "$0")"
echo "Instalando (solo la primera vez)..."
pip3 install -r requirements.txt >/dev/null 2>&1
echo "Abriendo el programa en tu navegador..."
open http://localhost:8000
python3 banrol_motor.py --serve --port 8000
