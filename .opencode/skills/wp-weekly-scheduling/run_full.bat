@echo off
cd /d "%~dp0."
python scripts\generate_schedule.py --mode full --input-dir input --output-dir output
pause
