@echo off
copy /y "E:\src\github\ExcelSheetMaster\dist\ESheetSearchMaster.exe" "E:\FW2\tools\ExcelSheetMaster\ESheetSearchMaster.exe"
cd /d "E:\FW2\tools\ExcelSheetMaster"
explorer "E:\FW2\tools\ExcelSheetMaster"
TortoiseProc.exe /command:commit /path:. /closeonend:0
