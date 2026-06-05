@echo off
title Consolidar PowerPoint por tienda

if "%~1"=="" (
    set /p CARPETA=Pegue la ruta de la carpeta a consolidar: 
) else (
    set "CARPETA=%~1"
)

cd /d "C:\1\OneDrive - Eficacia\KENVUE - Front\EXHIBICIONES ACUMULADAS Y ELEMENTOS\2026\Scripts"

python "consolidar_ppt_por_tienda.py" "%CARPETA%"

pause
