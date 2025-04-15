@echo off
setlocal
cd /d C:\Users\leonardo.fragoso\Desktop\Projetos\dash-burgetXLogComexXComercial

REM 1 - Ativar a venv
call .\venv\Scripts\activate.bat

REM 2 - Executar main.py que está na pasta "mes atual"
echo ===============================
echo Executando main.py (mes atual)...
echo ===============================
python "main.py"
if errorlevel 1 (
    echo ERRO: main.py encontrou erros.
    pause
    exit /b %errorlevel%
)

REM 3 - Executar atualizar_systracker.py
echo ===============================
echo Executando atualizar_systracker.py...
echo ===============================
python atualizar_systracker.py
if errorlevel 1 (
    echo ERRO: atualizar_systracker.py encontrou erros.
    pause
    exit /b %errorlevel%
)

REM 4 - Executar app.py
echo ===============================
echo Executando app.py...
echo ===============================
python app.py
if errorlevel 1 (
    echo ERRO: app.py encontrou erros.
    pause
    exit /b %errorlevel%
)

echo ===============================
echo Execução concluída com sucesso!
echo ===============================
pause
