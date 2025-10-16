@echo off
chcp 65001 > nul
title Executando Processo de RPA Completo

echo.
echo --- INICIANDO ETAPA 1: RPA DOWNLOADER ---
python RPA_downloader.py

echo.
echo --- INICIANDO ETAPA 2: CONVERSOR ---
python converter.py

echo.
echo --- INICIANDO ETAPA 3: VOLUMETRIA E NOTIFICACAO ---
python volumetria.py  :: Supondo que o envio para o Teams esteja neste arquivo

echo.
echo Processo finalizado com sucesso.
pause