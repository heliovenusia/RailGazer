@echo off
echo Running FOIS RailGazer Automation...
docker run --rm -v "%cd%":/app railgazer
echo --------------------------------------------
echo DONE. Excel file saved in this folder.
pause
