@echo off
cd /d "C:\Users\User\Desktop\Telegram Automation"

set LOG_DIR=C:\Users\User\Desktop\Telegram Automation\logs
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
set LOG_FILE=%LOG_DIR%\%date:~-4,4%-%date:~-7,2%-%date:~0,2%.txt

echo =============================== >> "%LOG_FILE%"
echo %date% %time% - Starting scraper >> "%LOG_FILE%"
echo =============================== >> "%LOG_FILE%"

call ".venv\Scripts\activate.bat"

python csroi.py >> "%LOG_FILE%" 2>&1
echo %date% %time% - Scraper done >> "%LOG_FILE%"

python TelegramNotifier.py >> "%LOG_FILE%" 2>&1
echo %date% %time% - Notifier done >> "%LOG_FILE%"

echo. >> "%LOG_FILE%"
