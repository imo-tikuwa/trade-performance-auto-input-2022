@echo off

cd /D %~dp0
call .\venv\Scripts\activate.bat
python app.py --debug
call deactivate.bat
pause