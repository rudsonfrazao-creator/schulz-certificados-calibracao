@echo off
setlocal enableextensions
chcp 65001 >nul
cd /d "%~dp0"
title Schulz S.A. - Automacao de Certificados

echo ============================================================
echo   Schulz S.A. - Automacao de Certificados de Calibracao
echo ============================================================

REM ---------- 1. Checagem do Python 3.10+ ----------
where python >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado no PATH.
    echo Instale o Python 3.10 ou superior em https://www.python.org/downloads/
    echo e marque a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)
python -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)"
if errorlevel 1 (
    echo [ERRO] E necessario Python 3.10 ou superior. Versao encontrada:
    python --version
    pause
    exit /b 1
)
echo [OK] Python detectado:
python --version

REM ---------- 2. Ambiente virtual ----------
if not exist "venv\Scripts\activate.bat" (
    echo [INFO] Criando ambiente virtual...
    python -m venv venv
    if errorlevel 1 (
        echo [ERRO] Falha ao criar o ambiente virtual.
        pause
        exit /b 1
    )
)
call "venv\Scripts\activate.bat"
echo [OK] Ambiente virtual ativado.

REM ---------- 3. Dependencias ----------
if not exist "requirements.txt" (
    echo [ERRO] requirements.txt nao encontrado na raiz do projeto.
    pause
    exit /b 1
)
echo [INFO] Instalando dependencias...
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias. Verifique a conexao com a internet.
    pause
    exit /b 1
)

REM ---------- 4. Servicos: navegador + API ----------
set PORT=5000
echo [INFO] Abrindo o navegador em http://127.0.0.1:%PORT% ...
start "" cmd /c "timeout /t 4 /nobreak >nul & start http://127.0.0.1:%PORT%"
echo [INFO] Iniciando API backend (Ctrl+C para encerrar)...
python app.py

pause
endlocal
