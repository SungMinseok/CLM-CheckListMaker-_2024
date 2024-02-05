@echo off
set "source=C:\Users\mssung\OneDrive\파이썬프로젝트\CLM(CheckListMaker)"
set "destination=C:\Users\mssung\OneDrive - Webzen Inc\R2M도구\CLM(CheckListMaker)"

xcopy "%source%" "%destination%" /E /EXCLUDE:%source%\exclude.txt
