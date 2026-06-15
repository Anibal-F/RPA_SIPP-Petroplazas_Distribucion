#!/bin/bash
echo "============================================"
echo " RPA SIPP Petroplazas — Instalacion macOS/Linux"
echo "============================================"
echo

# Verificar Python
if ! command -v python3 &>/dev/null; then
    echo "[ERROR] python3 no encontrado. Instálalo desde https://python.org"
    exit 1
fi

echo "[1/3] Instalando dependencias Python..."
pip3 install -r requirements.txt || { echo "[ERROR] Falló pip install"; exit 1; }

echo
echo "[2/3] Descargando navegador Chromium para Playwright..."
playwright install chromium || { echo "[ERROR] Falló playwright install"; exit 1; }

echo
echo "[3/3] Instalación completa."
echo
echo "Para ejecutar la aplicación:"
echo "    python3 main.py"
echo
