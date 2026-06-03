@echo off
chcp 65001 >nul
echo ============================================
echo  Сборка Читательского дневника
echo ============================================
echo.

echo [1/3] Установка зависимостей...
py -3.12 -m pip install -r requirements.txt
if errorlevel 1 (
    echo ОШИБКА: не удалось установить зависимости
    pause
    exit /b 1
)

echo.
echo [2/3] Сборка .exe через PyInstaller...
py -3.12 -m PyInstaller --clean --noconfirm ChitatelDnevnik.spec
if errorlevel 1 (
    echo ОШИБКА: сборка завершилась с ошибкой
    pause
    exit /b 1
)

echo.
echo [3/3] Копирование файлов в release\...
if not exist release mkdir release

copy /Y dist\ChitatelDnevnik.exe release\ChitatelDnevnik.exe
copy /Y init_db.sql release\init_db.sql
copy /Y README.md release\README.md

if not exist release\.env (
    copy /Y .env.example release\.env
    echo Создан release\.env из шаблона — укажите ваш пароль MySQL
)

echo.
echo ============================================
echo  Готово! Файлы в папке release\
echo  Не забудьте указать пароль в release\.env
echo ============================================
pause
