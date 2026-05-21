@echo off
cd /d "%~dp0"

if exist venv (
    echo Removing old venv...
    rmdir /s /q venv
)

python -m venv venv
call venv\Scripts\activate

python -m pip install --upgrade pip
pip install -r requirements.txt

python app.py
pause
