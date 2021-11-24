@echo off
echo Checking for Visio Installation
cd %SystemDrive%\users
dir /s completedoffice.txt 
if not errorlevel 1 goto end
echo Running installer
\\net_share\Setup.exe /configure \\net_share\configuration.xml
echo Office Install Complete > %SystemDrive%\users\completedoffice.txt
:end
exit
