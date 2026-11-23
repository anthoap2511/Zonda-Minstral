@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
 echo No se encontro Python. Instala Python 3.11 o superior desde https://www.python.org/downloads/
 pause
 exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
 echo Creando entorno virtual...
 py -m venv .venv
)
echo Instalando o verificando dependencias...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
 echo Hubo un error instalando dependencias.
 pause
 exit /b 1
)
echo Iniciando ZONDA - MISTRAL...
start "" http://127.0.0.1:5000
".venv\Scripts\python.exe" app.py
pause
