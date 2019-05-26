@echo off
REM echo %1 %2 %3
REM Pause

:start

cls

REM echo.Please input valid excel file without -ext,like as [test]

REM set /p input_source=

REM REM if not exist %input_source% echo.aaaa && goto start

REM goto go

REM :go
REM python do.py %input_source%
python do.py

pause