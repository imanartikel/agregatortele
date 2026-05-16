@echo off
:loop
echo [%date% %time%] Starting Telegram Bot...
python files/bot.py
echo [%date% %time%] Bot crashed or stopped. Restarting in 5 seconds...
timeout /t 5
goto loop
