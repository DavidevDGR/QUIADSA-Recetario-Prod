@echo off
REM ============================================================
REM  Compila AppFabricacion.exe con PyInstaller.
REM  Ejecutar este script DENTRO de la carpeta del proyecto,
REM  en un símbolo de sistema (cmd) o PowerShell, en Windows.
REM ============================================================
REM  NOTA: se usa "python -m pip" / "python -m PyInstaller" en lugar de
REM  "pip" / "pyinstaller" directamente. Así funciona aunque la carpeta
REM  "Scripts" de Python no esté en el PATH del sistema (causa más
REM  habitual del error "pyinstaller no se reconoce como un comando").
REM ============================================================

echo.
echo === 0/3: Comprobando que Python está accesible ===
python --version
if errorlevel 1 (
    echo.
    echo ERROR: "python" no se reconoce. Instala Python desde python.org
    echo y asegurate de marcar la casilla "Add python.exe to PATH" durante
    echo la instalacion.
    pause
    exit /b 1
)

echo.
echo === 1/3: Instalando dependencias ===
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR instalando dependencias.
    pause
    exit /b 1
)

echo.
echo === 2/3: Limpiando compilaciones anteriores ===
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q AppFabricacion.spec 2>nul

echo.
echo === 3/3: Generando el ejecutable (dist\AppFabricacion.exe) ===
REM --onefile     -> un unico .exe (mas comodo para llevar a otro PC)
REM --windowed    -> sin consola negra de fondo (app de escritorio)
REM (ya NO se excluye pyodbc: hace falta porque MOCK_MODE=False usa la
REM  conexion real a Solmicro. Si en el futuro vuelves a MOCK_MODE=True y
REM  quieres un .exe mas ligero, puedes anadir de nuevo
REM  "--exclude-module pyodbc" a la linea de abajo.)
python -m PyInstaller --noconfirm --onefile --windowed --name AppFabricacion main.py

if errorlevel 1 (
    echo.
    echo ERROR generando el ejecutable. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  LISTO. El ejecutable esta en: dist\AppFabricacion.exe
echo  Para llevarlo a otro equipo, copia SOLO ese fichero .exe
echo  (no hace falta llevar Python ni las carpetas build/dist).
echo ============================================================
pause