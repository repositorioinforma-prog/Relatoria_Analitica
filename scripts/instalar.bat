@echo off
cd /d "%~dp0.."
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
    if errorlevel 1 goto erro
)
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto erro
echo Instalacao concluida.
pause
exit /b 0
:erro
echo Falha na instalacao. Confira a mensagem acima e a instalacao do Python.
pause
exit /b 1
