@REM Program Section 7: Update PowerTerm 10 and PowerTerm 12 users to the updated Auto-Populating forms that work with those PowerTerm versions

	IF %PowerTermDirectoryFilePath%=="C:\Program Files\Ericom Software\PowerTerm" SET WmicPowerTermDirectoryFilePath='C:\\Program Files\\Ericom Software\\PowerTerm\\ptw32.exe'
	IF %PowerTermDirectoryFilePath%=="C:\Program Files\Ericom Software\PowerTerm10" SET WmicPowerTermDirectoryFilePath='C:\\Program Files\\Ericom Software\\PowerTerm10\\ptw32.exe'
	IF %PowerTermDirectoryFilePath%=="C:\Program Files (x86)\Ericom Software\PowerTerm" SET WmicPowerTermDirectoryFilePath='C:\\Program Files (x86)\\Ericom Software\\PowerTerm\\ptw32.exe'
	IF %PowerTermDirectoryFilePath%=="C:\Program Files (x86)\Ericom Software\PowerTerm10" SET WmicPowerTermDirectoryFilePath='C:\\Program Files (x86)\\Ericom Software\\PowerTerm10\\ptw32.exe'

	SETLOCAL enableDelayedExpansion
	FOR /F "tokens=2 delims==" %%I IN (
		'wmic datafile where "name=%WmicPowerTermDirectoryFilePath%" get version /format:list'
	) DO FOR /F "delims=" %%A IN ("%%I") DO SET "RESULT=%%A"
	
	@REM %Result% will give you the result of the above for loop. We're setting the PowerTerm version number to the first two digits of that.
	@REM So if the '%Result%' was 12.5.0.2944, then it will end up being '12' instead. %RESULT:~0,2% means the first two digits of the result variable.
	SET PowerTermVersionNumber=%RESULT:~0,2%

	ECHO The PowerTerm Version number is %PowerTermVersionNumber%
	
	@REM Copy the updated automated forms to PowerTerm 10 and PowerTerm 12 users folders.
	IF %PowerTermVersionNumber%==12 %systemroot%\System32\ROBOCOPY "\\qdcns0002\tmp_data$\tmpdept\Knowledge_Base\Technical-Repositories\Excel Macro Workbooks\Updated Forms" %PowerTermDirectoryFilePath% /XX
	IF %PowerTermVersionNumber%==10 %systemroot%\System32\ROBOCOPY "\\qdcns0002\tmp_data$\tmpdept\Knowledge_Base\Technical-Repositories\Excel Macro Workbooks\Updated Forms" %PowerTermDirectoryFilePath% /XX