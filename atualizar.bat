@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0publicar.ps1"
echo.
echo ============================================
echo Pressione qualquer tecla para fechar esta janela.
echo ============================================
pause >nul
