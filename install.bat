@echo off
echo Checking for Visio Installation
cd %SystemDrive%\users
dir /s completedoffice.txt 
if not errorlevel 1 goto end
echo Running installer
\\srv-db1\Config.1C$\Office2019\Setup.exe /configure \\srv-db1\Config.1C$\Office2019\configuration.xml
echo Office Install Complete > %SystemDrive%\users\completedoffice.txt
:end
exit