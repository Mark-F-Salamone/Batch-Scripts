@ECHO off

@REM Test that the user has the proper permissions to update files in their PowerTerm folder

SET testdir=%UserProfile%\Desktop
SET myguid={A4E30755-FE04-4ab7-BD7F-E006E37B7BF7}.tmp
SET waccess=0

ECHO.> %testdir%\%myguid%&&(SET waccess=1&del %testdir%\%myguid%)
ECHO write access=%waccess%

IF %waccess%==1 (
	ECHO You've got Read/Write permissions on %testdir%. 
)

IF %waccess%==0 (

	ECHO You do not have Read/Write permissions on %testdir%.
	)