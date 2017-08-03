@echo off
copy /y "%~dp0"TABCTL32.OCX c:\windows\system32
copy /y "%~dp0"MSCOMCTL.OCX c:\windows\system32
copy /y "%~dp0"mscomct2.ocx c:\windows\system32
regsvr32 c:\windows\system32\TABCTL32.OCX
regsvr32 c:\windows\system32\MSCOMCTL.OCX
regsvr32 c:\windows\system32\mscomct2.ocx
echo. ÍË³ö
pause