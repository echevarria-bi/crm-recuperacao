@echo off
title Atualizar CRM Recuperação
echo ========================================
echo   Atualizador CRM Recuperação
echo ========================================
echo.

set "BASEDIR=C:\Users\mar\OneDrive - SPADER DISTRIBUIDORA DE ALIMENTOS L\Area de Trabalho"
set "CAMDIR=%BASEDIR%\dashboards\crm-recuperacao"
set "SCRIPTDIR=%~dp0"

echo [1/4] Extraindo dados da base...
node "%SCRIPTDIR%extrair_crm.js"
if %ERRORLEVEL% neq 0 (
    echo ERRO ao extrair dados!
    pause
    exit /b 1
)
echo OK!
echo.

echo [2/4] Enviando para GitHub...
git add data.json
git commit -m "feat: atualizacao crm recuperacao %date% %time%" 2>nul
git push
echo.

echo [3/4] Resumo:
node -e "var d=require('%CAMDIR%\\data.json');var fc=d.data.faturamentoClientes;var mt={};var mc={};fc.forEach(function(c){if(!mt[c.mes])mt[c.mes]=0;if(!mc[c.mes])mc[c.mes]={};mt[c.mes]+=c.valor;mc[c.mes][c.cliente]=1;});console.log('  Abril: R\$ '+(mt.Abril||0).toFixed(2)+' | '+Object.keys(mc.Abril||{}).length+' clientes');console.log('  Maio: R\$ '+(mt.Maio||0).toFixed(2)+' | '+Object.keys(mc.Maio||{}).length+' clientes');console.log('  Junho: R\$ '+(mt.Junho||0).toFixed(2)+' | '+Object.keys(mc.Junho||{}).length+' clientes');console.log('  Julho: R\$ '+(mt.Julho||0).toFixed(2)+' | '+Object.keys(mc.Julho||{}).length+' clientes');var t=fc.reduce(function(s,c){return s+c.valor;},0);console.log('  Total acumulado: R\$ '+t.toFixed(2));console.log('  Total clientes: '+Object.keys(d.data.clientesDetalhes).length);"
echo.

echo [4/4] Dashboard atualizado!
pause
