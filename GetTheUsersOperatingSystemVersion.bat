@ECHO off

@REM Find out if the computer is a 32BIT or a 64BIT Operating System

%systemroot%\System32\REG Query "HKLM\Hardware\Description\System\CentralProcessor\0" | %systemroot%\System32\FIND /i "x86" > NUL && SET OS=32BIT || SET OS=64BIT

IF %OS%==32BIT ECHO This is a 32bit operating system
IF %OS%==64BIT ECHO This is a 64bit operating system
@ECHO: