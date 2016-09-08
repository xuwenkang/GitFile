@echo off
set file = 1
set path=%cd%\SYSINFO.OCX
echo %path%
copy "%path%" "C:\WINDOWS\system"
regsvr32 /u /s sysinfo.ocx