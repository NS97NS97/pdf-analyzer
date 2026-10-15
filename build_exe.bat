@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Сборка АнализаторPDF.exe

echo ============================================================
echo   Сборка автономного АнализаторPDF.exe (Windows)
echo ============================================================
echo.

echo [1/3] Проверяю и устанавливаю зависимости...
py -m pip install --upgrade pyinstaller flask pypdf openpyxl
if errorlevel 1 (
    echo.
    echo [Ошибка] Не удалось установить зависимости. Проверьте, что установлен Python.
    pause
    exit /b 1
)

echo.
echo [2/3] Собираю .exe (это может занять 1-3 минуты)...
py -m PyInstaller --noconfirm --clean --onefile ^
    --name "АнализаторPDF" ^
    --add-data "templates;templates" ^
    --add-data "static;static" ^
    --collect-submodules pypdf ^
    --exclude-module numpy ^
    --exclude-module pandas ^
    --exclude-module scipy ^
    --exclude-module matplotlib ^
    --exclude-module PIL ^
    --exclude-module IPython ^
    --exclude-module pytest ^
    --exclude-module notebook ^
    --exclude-module sqlite3 ^
    app.py
if errorlevel 1 (
    echo.
    echo [Ошибка] Сборка не удалась. Смотрите сообщения выше.
    pause
    exit /b 1
)

echo.
echo [3/3] Готово!
echo.
echo   Программа:  %~dp0dist\АнализаторPDF.exe
echo.
echo Это автономный файл — его можно скопировать на любой компьютер
echo с 64-разрядной Windows и запускать двойным щелчком (Python не нужен).
echo.
pause
