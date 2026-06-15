@echo off
echo ============================================
echo  RPA SIPP Petroplazas — Instalacion Windows
echo ============================================
echo.

:: Verificar Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no encontrado. Descargalo desde https://python.org
    pause
    exit /b 1
)

echo [1/3] Instalando dependencias Python...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Fallo al instalar dependencias.
    pause
    exit /b 1
)

echo.
echo [2/3] Descargando navegador Chromium para Playwright...
playwright install chromium
if %errorlevel% neq 0 (
    echo [ERROR] Fallo al instalar Playwright.
    pause
    exit /b 1
)

echo.
echo [3/3] Instalacion completa.
echo.
echo Para ejecutar la aplicacion:
echo     python main.py
echo.
pause
