@echo off
cd /d "%~dp0"
pip install -r ../Portal-de-boletos/requirements.txt
python "all_in_one.py"
exit
