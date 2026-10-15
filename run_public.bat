@echo off
chcp 65001 >nul
title Анализатор документов — публичный доступ

echo ============================================================
echo   Запускаю сайт на этом компьютере и публичную ссылку.
echo   Файлы остаются только у вас и НЕ копируются в интернет.
echo ============================================================
echo.

where cloudflared >nul 2>nul
if errorlevel 1 (
    echo [Ошибка] Не найден cloudflared.
    echo Установите его: https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/
    echo Либо используйте ngrok:  ngrok http 5000
    echo.
    pause
    exit /b 1
)

echo [1/2] Запускаю сайт на http://127.0.0.1:5000 ...
start "Анализатор документов" py app.py 5000

echo [2/2] Открываю публичную ссылку (она появится ниже) ...
timeout /t 3 >nul
cloudflared tunnel --url http://127.0.0.1:5000

pause
