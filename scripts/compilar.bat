@echo off
cd /d "%~dp0.."
if not exist ".venv\Scripts\python.exe" (
    echo Execute scripts\instalar.bat primeiro.
    pause
    exit /b 1
)
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
if errorlevel 1 goto erro
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --distpath dist --workpath build "packaging\Processamentos Relatoria.spec"
if errorlevel 1 goto erro
echo Aplicativo gerado em dist\Processamentos Relatoria.
pause
exit /b 0
:erro
echo Falha na compilacao. Confira a mensagem acima.
pause
exit /b 1
