@echo off
cd /d "d:\CLG\excel automation"
.venv\Scripts\activate
echo ai_student_details.xlsx | .venv\Scripts\python.exe update_profiles.py
pause