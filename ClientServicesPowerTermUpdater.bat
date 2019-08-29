@ECHO off 
	
@REM Program Section 1: Determine the correct directory file paths starting with the user's correct PowerTerm directory

	@REM Section 1.0: Declare Variables
	
	SET CurrentPowerTermUpdaterVersion=ClientServicesPowerTermUpdater-1.0.4.lnk
	SET PreviousPowerTermUpdaterVersion=ClientServicesPowerTermUpdater-1.0.3.lnk
	SET NoNeedToUpdatePtpFiles=False
	
	@REM Section 1.1: Find out if the computer is a 32BIT or a 64BIT Operating System
	%systemroot%\System32\REG Query "HKLM\Hardware\Description\System\CentralProcessor\0" | %systemroot%\System32\FIND /i "x86" > NUL && SET OS=32BIT || SET OS=64BIT
	
	IF %OS%==32BIT ECHO This is a 32bit operating system
	IF %OS%==64BIT ECHO This is a 64bit operating system
	@ECHO:
	
	@REM Section 1.2: Assign the user's correct PowerTerm Directory File Path based on their Operating System of either 32 Bit or 64 Bit
	IF %OS%==64BIT (
		@REM In the below code, C:\Program Files\Ericom Software\PowerTerm10 is the 4th choice and C:\Program Files (x86)\Ericom Software\PowerTerm is the 1st choice for 64 Bit computers.
		IF EXIST "C:\Program Files\Ericom Software\PowerTerm10\ptw32.exe" SET "PowerTermDirectoryFilePath="C:\Program Files\Ericom Software\PowerTerm10""
		IF EXIST "C:\Program Files (x86)\Ericom Software\PowerTerm10\ptw32.exe" SET "PowerTermDirectoryFilePath="C:\Program Files ^(x86^)\Ericom Software\PowerTerm10""
		IF EXIST "C:\Program Files\Ericom Software\PowerTerm\ptw32.exe" SET "PowerTermDirectoryFilePath="C:\Program Files\Ericom Software\PowerTerm""
		IF EXIST "C:\Program Files (x86)\Ericom Software\PowerTerm\ptw32.exe" SET "PowerTermDirectoryFilePath="C:\Program Files ^(x86^)\Ericom Software\PowerTerm""
	)
	
	IF %OS%==32BIT (
		@REM In the below code, C:\Program Files\Ericom Software\PowerTerm10\ is the 2nd choice for 32 Bit computers.
		IF EXIST "C:\Program Files\Ericom Software\PowerTerm10\ptw32.exe" SET "PowerTermDirectoryFilePath="C:\Program Files\Ericom Software\PowerTerm10""
		IF EXIST "C:\Program Files\Ericom Software\PowerTerm\ptw32.exe" SET "PowerTermDirectoryFilePath="C:\Program Files\Ericom Software\PowerTerm""
	)
	
	ECHO PowerTerm Directory: %PowerTermDirectoryFilePath%
	@ECHO:
	
	@REM Section 1.3: Assign the user's correct Terminals Directory File Path based off of the user's PowerTermDirectoryFilePath
	IF %PowerTermDirectoryFilePath%=="C:\Program Files\Ericom Software\PowerTerm" SET TerminalsSharedDriveDirectoryFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\Terminals
	IF %PowerTermDirectoryFilePath%=="C:\Program Files\Ericom Software\PowerTerm10" SET TerminalsSharedDriveDirectoryFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\Terminals_10
	IF %PowerTermDirectoryFilePath%=="C:\Program Files (x86)\Ericom Software\PowerTerm" SET TerminalsSharedDriveDirectoryFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\Terminals_x86
	IF %PowerTermDirectoryFilePath%=="C:\Program Files (x86)\Ericom Software\PowerTerm10" SET TerminalsSharedDriveDirectoryFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\Terminals10_x86
	
	ECHO Terminals Directory: %TerminalsSharedDriveDirectoryFilePath%
	@ECHO:
	
	@REM Section 1.4: Assign the user's correct QLS Shortcuts Directory File Path based off of the user's PowerTermDirectoryFilePath
	IF %PowerTermDirectoryFilePath%=="C:\Program Files\Ericom Software\PowerTerm" SET QLSShortcutsSharedDriveDirectoryFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\QLS_Shortcuts
	IF %PowerTermDirectoryFilePath%=="C:\Program Files\Ericom Software\PowerTerm10" SET QLSShortcutsSharedDriveDirectoryFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\QLS_Shortcuts_10
	IF %PowerTermDirectoryFilePath%=="C:\Program Files (x86)\Ericom Software\PowerTerm" SET QLSShortcutsSharedDriveDirectoryFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\QLS_Shortcuts_x86
	IF %PowerTermDirectoryFilePath%=="C:\Program Files (x86)\Ericom Software\PowerTerm10" SET QLSShortcutsSharedDriveDirectoryFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\QLS_Shortcuts10_x86
	
	ECHO QLSShortcuts Directory: %QLSShortcutsSharedDriveDirectoryFilePath%
	@ECHO:

	@REM Section 1.5: Test that the user has the proper permissions to update files in their PowerTerm folder
	
	SET testdir=%PowerTermDirectoryFilePath%
	SET myguid={A4E30755-FE04-4ab7-BD7F-E006E37B7BF7}.tmp
	SET waccess=0
	
	ECHO.> %testdir%\%myguid%&&(SET waccess=1&del %testdir%\%myguid%)
	ECHO write access=%waccess%
	
	IF %waccess%==1 (
		ECHO You've got the right permissions on %PowerTermDirectoryFilePath%. The installation can continue!
	)

	IF %waccess%==0 (

		START \\QDCNS0002\TMP_Data$\TMPDept\Knowledge_Base\Programs\PowerTermFolderPermissions.exe
		EXIT
		)
		
	@REM Section 1.6: Set the WmicPowerTermDirectoryFilePath in order to find the user's PowerTerm version later in the program
	
	IF %PowerTermDirectoryFilePath%=="C:\Program Files\Ericom Software\PowerTerm" SET WmicPowerTermDirectoryFilePath='C:\\Program Files\\Ericom Software\\PowerTerm\\ptw32.exe'
	IF %PowerTermDirectoryFilePath%=="C:\Program Files\Ericom Software\PowerTerm10" SET WmicPowerTermDirectoryFilePath='C:\\Program Files\\Ericom Software\\PowerTerm10\\ptw32.exe'
	IF %PowerTermDirectoryFilePath%=="C:\Program Files (x86)\Ericom Software\PowerTerm" SET WmicPowerTermDirectoryFilePath='C:\\Program Files (x86)\\Ericom Software\\PowerTerm\\ptw32.exe'
	IF %PowerTermDirectoryFilePath%=="C:\Program Files (x86)\Ericom Software\PowerTerm10" SET WmicPowerTermDirectoryFilePath='C:\\Program Files (x86)\\Ericom Software\\PowerTerm10\\ptw32.exe'
		
@REM Program Section 2: Kill PowerTerm Processes

	ECHO Closing all active PowerTerm windows so that the Power Pad files can be updated.
	@ECHO:
	%systemroot%\System32\taskkill /im ptw32.exe /f
	
@REM Program Section 3: Update The PowerTerm Icons, Shortcuts, and Terminal Setup Files
	
	@REM If the ClientServicesPowerTermUpdater.lnk file does not exist in the user's StartUp folder 
	@REM Then that means the user hasn't had this program run and we need to ensure there files are correct.
	
	IF NOT EXIST "%AppData%\Microsoft\Windows\Start Menu\Programs\Startup\%CurrentPowerTermUpdaterVersion%" (
		
		@REM Program Section 3.1: Update the user's icons, QLS Shortcuts, and bring the ClientServicesPowerTermUpdater.lnk file to the user's StartUp folder.
		@ECHO This is an initial install. Updating your icons, QLS Shortcuts, and %CurrentPowerTermUpdaterVersion% file.
		@ECHO:
		
		@REM Update the name of the shortcut to whatever the programmer sets as the CurrentPowerTermVersion in the variable above.
		IF EXIST \\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\Batch\Batch_Shortcuts\%PreviousPowerTermUpdaterVersion% (
			RENAME \\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\Batch\Batch_Shortcuts\%PreviousPowerTermUpdaterVersion% %CurrentPowerTermUpdaterVersion%
		)
		
		%systemroot%\System32\XCOPY /y \\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\Batch\Batch_Shortcuts\%CurrentPowerTermUpdaterVersion% "%AppData%\Microsoft\Windows\Start Menu\Programs\Startup"
		%systemroot%\System32\ROBOCOPY /MIR \\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\Icons %PowerTermDirectoryFilePath%\icons
		%systemroot%\System32\ROBOCOPY /MIR %QLSShortcutsSharedDriveDirectoryFilePath% %UserProfile%\Desktop\QLSShortcuts
		
		@REM Users with Windows 10 have a different file path for the Desktop that they view. This ensures the QLS Shortcuts directory gets on that desktop.
		
		IF EXIST "%UserProfile%\OneDrive - Quest Diagnostics\Desktop" (
			ECHO There is a OneDrive Desktop on your computer. Copying a QLSShortcuts folder to it now.
			%systemroot%\System32\ROBOCOPY /MIR %QLSShortcutsSharedDriveDirectoryFilePath% "%UserProfile%\OneDrive - Quest Diagnostics\Desktop\QLSShortcuts"
		)
		
		@REM Program Section 3.2: If the user already has a local folder, update their psl, ptc, and pts files. Do not update their ptk or Hot Key files. .ptp files are updated in Section 5
		IF EXIST %PowerTermDirectoryFilePath%\local\NUL (
			@ECHO:
			ECHO Updating your .psl, .ptc, and .pts PowerTerm Terminal Setup files.
			@ECHO:
			FOR /R %TerminalsSharedDriveDirectoryFilePath% %%f IN (*.psl) DO COPY "%%f" %PowerTermDirectoryFilePath%\local
			FOR /R %TerminalsSharedDriveDirectoryFilePath% %%f IN (*.ptc) DO COPY "%%f" %PowerTermDirectoryFilePath%\local
			FOR /R %TerminalsSharedDriveDirectoryFilePath% %%f IN (*.pts) DO COPY "%%f" %PowerTermDirectoryFilePath%\local
		)
		
		@REM Program Section 3.3 If the user doesn't have a local folder, just bring over a mirror copy from the appropriate Terminals Shared Drive Directory.
		@ECHO:
		ECHO Creating a new local folder for the user complete with all necessary PowerTerm Terminal Setup files.
		@ECHO:
		
		@REM This will remove the file called "local" and add the appropriate "local" directory to the user's PowerTerm folder
		IF NOT EXIST %PowerTermDirectoryFilePath%\local\NUL (
			
			ECHO The local directory doesn't exist
			DEL %PowerTermDirectoryFilePath%\local /Q
			%systemroot%\System32\ROBOCOPY /MIR %TerminalsSharedDriveDirectoryFilePath% %PowerTermDirectoryFilePath%\local
			SET NoNeedToUpdatePtpFiles=True
		)
		
		@REM Program Section 3.4: Update the user's PT-Comm file

		IF %PowerTermDirectoryFilePath%=="C:\Program Files\Ericom Software\PowerTerm" SET PTCommFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\PtCommFiles\ptcomm.ini
		IF %PowerTermDirectoryFilePath%=="C:\Program Files\Ericom Software\PowerTerm10" SET PTCommFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\PtCommFiles\ptcomm_10.ini
		IF %PowerTermDirectoryFilePath%=="C:\Program Files (x86)\Ericom Software\PowerTerm" SET PTCommFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\PtCommFiles\ptcomm_x86.ini
		IF %PowerTermDirectoryFilePath%=="C:\Program Files (x86)\Ericom Software\PowerTerm10" SET PTCommFilePath=\\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\PowerTerm\PtCommFiles\ptcomm10_x86.ini
		
		ECHO The user's PTComm File Path is: %PTCommFilePath%
		
		%systemroot%\System32\XCOPY /y %PTCommFilePath% %PowerTermDirectoryFilePath%\ptcomm.ini
	)
	
@REM Program Section 4: Always copy the ClientServicesPowerTermUpdater.lnk file to the user's computer just in case the .lnk file location needs to be updated or other reasons.
	@ECHO:
	ECHO Copying the %CurrentPowerTermUpdaterVersion% file to your StartUp folder.
	@ECHO:
	%systemroot%\System32\XCOPY /y \\qdcns0002\tmp_data$\TMPDept\Knowledge_Base\Batch\Batch_Shortcuts\%CurrentPowerTermUpdaterVersion% "%AppData%\Microsoft\Windows\Start Menu\Programs\Startup"
	
@REM Program Section 5: Update the user's Power Pad files unless a brand new local folder has already been brought over
	@ECHO:
	
	IF %NoNeedToUpdatePtpFiles%==False (
		ECHO Updating your Power Pad .ptp files.
		@ECHO:
		FOR /R %TerminalsSharedDriveDirectoryFilePath% %%f IN (*.ptp) DO COPY "%%f" %PowerTermDirectoryFilePath%\local
		@ECHO:
	)
	
@REM Program Section 6: Update the Power Pad Resources. Including but not limited to: Emails, Forms, References, and URLs
	@ECHO:
	ECHO Updating your Power Pad Resources (Emails, Forms, References, URLs, etc...)
	@ECHO:
	%systemroot%\System32\ROBOCOPY /XX /MIR \\QDCNS0002\TMP_DATA$\TMPDEPT\Knowledge_Base\Links %PowerTermDirectoryFilePath%
	
@REM Program Section 7: Update PowerTerm 10 and PowerTerm 12 users to the updated Auto-Populating forms that work with those PowerTerm versions

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
	
@REM Program Section 8: Delete any unnecessary files

	IF EXIST %AppData%\"Microsoft\Windows\Start Menu\Programs\Startup\%PreviousPowerTermUpdaterVersion%" DEL %AppData%\"Microsoft\Windows\Start Menu\Programs\Startup\%PreviousPowerTermUpdaterVersion%"
	IF EXIST %AppData%\"Microsoft\Windows\Start Menu\Programs\Startup\Update_Pad.lnk" DEL %AppData%\"Microsoft\Windows\Start Menu\Programs\Startup\Update_Pad.lnk"
	IF EXIST %UserProfile%\Desktop\QLS_Shortcuts RMDIR %UserProfile%\Desktop\QLS_Shortcuts /Q /S
	IF EXIST %UserProfile%\Desktop\QLS_Shortcuts_10 RMDIR %UserProfile%\Desktop\QLS_Shortcuts_10 /Q /S
	IF EXIST %UserProfile%\Desktop\QLS_Shortcuts_x86 RMDIR %UserProfile%\Desktop\QLS_Shortcuts_x86 /Q /S
	IF EXIST %UserProfile%\Desktop\QLS_Shortcuts10_x86 RMDIR %UserProfile%\Desktop\QLS_Shortcuts10_x86 /Q /S

	@REM Program Footer

@ECHO:
ECHO *************************************************************************
@ECHO:
ECHO ---- Program last updated by Mark F. Salamone on 04/10/2019 1:37 PM ----
@ECHO:
ECHO *************************************************************************
@ECHO: