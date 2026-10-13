@echo off
cd /d "%~dp0."
python scripts\generate_schedule.py --mode append --input-dir input --output-dir output
pause
