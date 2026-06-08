@echo off
set GARAMUT_APP_PASSWORD=%GARAMUT_APP_PASSWORD%
set GITHUB_TOKEN=%GITHUB_TOKEN%
"C:\Users\NormanReed\AppData\Local\Programs\Python\Python313\python.exe" ^
    "C:\Users\NormanReed\Documents\GitHub\thegaramut\process_garamut.py" ^
    >> "C:\Users\NormanReed\Documents\GitHub\thegaramut\garamut_task.log" 2>&1
