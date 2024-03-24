:: @echo off

if not exist ".venv" (

    echo Creating virtual environment..
    python -m venv .venv

    :: Activate the virtual environment
    call .venv\Scripts\activate.bat

    echo Installing dependencies..
    pip install --upgrade pip
    pip install -r requirements.txt
)


echo Using virtual environment..
call .venv\Scripts\activate.bat


echo Launching create_RDC.py..
python create_RDC.py

pause
