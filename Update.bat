@echo off
echo Update process ...
timeout /t 2
del GlucoTT.exe
ren Update.exe GlucoTT.exe
start GlucoTT.exe
exit
