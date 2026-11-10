@echo off
rem Ejecuta Exportar-Pericia.ps1 con el usuario actual, sin admin.
rem -ExecutionPolicy Bypass aplica solo a este proceso (no cambia nada del equipo).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Exportar-Pericia.ps1" %*
echo.
pause
