@echo off

@set dbPath=YOUR_DB_PATH
@set root=YOUR_ROOT_PATH

cmd.exe

pushd %root%

if exist ./venv (
    .venv\Scripts\activate
    pip freeze > requirements.txt
) else (
    python3 -m venv .venv
    .venv\Scripts\activate
    pip install -r requirements.txt
)

python creon_datareader.py --dbPath

pause