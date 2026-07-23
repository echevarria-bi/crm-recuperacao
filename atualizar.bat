@echo off
title Atualizar CRM Recuperação
echo ========================================
echo   Atualizador CRM Recuperação
echo ========================================
echo.

set "BASEDIR=C:\Users\mar\OneDrive - SPADER DISTRIBUIDORA DE ALIMENTOS L\Area de Trabalho"
set "CAMDIR=%BASEDIR%\dashboards\crm-recuperacao"
set "SCRIPTDIR=%~dp0"

echo [1/3] Extraindo dados da base...
node "%SCRIPTDIR%extrair_crm.js"
if %ERRORLEVEL% neq 0 (
    echo ERRO ao extrair dados!
    pause
    exit /b 1
)
echo OK!
echo.

echo [2/3] Enviando para GitHub...
git add data.json
git commit -m "feat: atualizacao crm recuperacao %date% %time%" 2>nul
git push
echo.

echo [3/3] Dashboard atualizado!
pause
