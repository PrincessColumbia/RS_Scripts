:@ECHO OFF
SET output=output.txt
IF EXIST "%output%" DEL "%output%"
FOR /f %%a IN (URLs.txt) DO (
    CALL :ping %%a

)
GOTO :EOF

:ping
ping -n 1 %1 | find "Approximate round trip" >NUL || ECHO %1>>"%output%"