@echo off
REM Gera o executavel ControleDespesas.exe na pasta "dist".
REM Requer Python 3.10+ instalado (https://www.python.org/downloads/ - marque "Add python.exe to PATH").
cd /d "%~dp0"
python -m venv .venv || goto erro
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt || goto erro
pyinstaller --noconfirm --clean --onefile --windowed --name ControleDespesas --icon icone.ico --add-data "icone.ico;." --add-data "icone.png;." controle_despesas.py || goto erro
echo.
echo Pronto! O executavel esta em: %~dp0dist\ControleDespesas.exe
pause
exit /b 0
:erro
echo.
echo Ocorreu um erro durante o build.
pause
exit /b 1
