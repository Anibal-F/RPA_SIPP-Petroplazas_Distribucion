@echo off
setlocal enabledelayedexpansion

:: Directorio donde vive este .bat (sin barra final)
set "APPDIR=%~dp0"
if "%APPDIR:~-1%"=="\" set "APPDIR=%APPDIR:~0,-1%"

echo ============================================
echo  RPA SIPP Petroplazas - Instalacion Windows
echo ============================================
echo.

:: ── 1. Verificar Python ──────────────────────────────────────────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no encontrado.
    echo         Descargalo desde https://python.org y asegurate de marcar
    echo         "Add Python to PATH" durante la instalacion.
    pause & exit /b 1
)

:: ── 2. Instalar dependencias ─────────────────────────────────────────────
echo [1/4] Instalando dependencias Python...
pip install -r "%APPDIR%\requirements.txt"
if %errorlevel% neq 0 (
    echo [ERROR] Fallo al instalar dependencias.
    pause & exit /b 1
)

:: ── 3. Playwright: descargar Chromium ────────────────────────────────────
echo.
echo [2/4] Descargando navegador Chromium para Playwright...
playwright install chromium
if %errorlevel% neq 0 (
    echo [ERROR] Fallo al instalar Playwright/Chromium.
    pause & exit /b 1
)

:: ── 4. Crear Logo_Petroil.ico desde PNG ─────────────────────────────────
echo.
echo [3/4] Generando icono .ico desde el logo...
python -c ^
  "from PIL import Image; " ^
  "img=Image.open(r'%APPDIR%\Logo_Petroil.png').convert('RGBA'); " ^
  "img.save(r'%APPDIR%\Logo_Petroil.ico', format='ICO', sizes=[(256,256),(128,128),(64,64),(32,32),(16,16)]); " ^
  "print('   Logo_Petroil.ico creado.')" 2>nul
if not exist "%APPDIR%\Logo_Petroil.ico" (
    echo    [AVISO] No se pudo crear el icono personalizado.
    echo            El acceso directo usara el icono de Python.
)

:: ── 5. Acceso directo en el escritorio ───────────────────────────────────
echo.
echo [4/4] Creando acceso directo en el escritorio...

:: Obtener ruta a pythonw.exe (sin ventana de consola) via Python
for /f "delims=" %%P in ('python -c "import sys,os; pw=os.path.join(os.path.dirname(sys.executable),'pythonw.exe'); print(pw if os.path.exists(pw) else sys.executable)"') do set "PYTHONW=%%P"

set "SHORTCUT=%USERPROFILE%\Desktop\RPA Petroplazas.lnk"
set "ICON=%APPDIR%\Logo_Petroil.ico"

:: Escribir script PowerShell temporal
set "PSFILE=%TEMP%\rpa_shortcut_%RANDOM%.ps1"
> "%PSFILE%"  echo $ws = New-Object -ComObject WScript.Shell
>> "%PSFILE%" echo $sc = $ws.CreateShortcut('%SHORTCUT%')
>> "%PSFILE%" echo $sc.TargetPath      = '%PYTHONW%'
>> "%PSFILE%" echo $sc.Arguments       = '"%APPDIR%\main.py"'
>> "%PSFILE%" echo $sc.WorkingDirectory = '%APPDIR%'
>> "%PSFILE%" echo $sc.IconLocation    = '%ICON%'
>> "%PSFILE%" echo $sc.Description     = 'RPA Recepcion de Facturas SIPP Petroplazas'
>> "%PSFILE%" echo $sc.Save()

powershell -NoProfile -ExecutionPolicy Bypass -File "%PSFILE%"
del "%PSFILE%" 2>nul

if exist "%SHORTCUT%" (
    echo    Acceso directo creado: RPA Petroplazas en escritorio.
) else (
    echo    [AVISO] No se pudo crear el acceso directo automaticamente.
    echo            Puedes crear uno manualmente apuntando a:
    echo            "%PYTHONW%" "%APPDIR%\main.py"
)

:: ── Listo ─────────────────────────────────────────────────────────────────
echo.
echo ============================================
echo  Instalacion completa.
echo  Abre la app con el icono "RPA Petroplazas"
echo  en tu escritorio, o ejecuta:
echo    python "%APPDIR%\main.py"
echo ============================================
echo.
pause
