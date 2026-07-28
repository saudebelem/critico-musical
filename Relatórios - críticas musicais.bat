@echo off
title Relatorios - Criticas Musicais

python --version >nul 2>&1
if errorlevel 1 goto NOPYTHON

python -c "import docx" >nul 2>&1
if errorlevel 1 python -m pip install python-docx

for %%F in ("%~dp0Relat*.py") do set "PY_SCRIPT=%%F"
python "%PY_SCRIPT%" %*
goto END

:NOPYTHON
echo [ERRO] O Python nao foi encontrado neste computador.
echo Baixe e instale o Python em https://www.python.org/downloads/
echo Marque a opcao "Add Python to PATH" durante a instalacao.
echo.

:END
pause
