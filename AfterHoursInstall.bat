@echo off
:: copy VbaProject to Outlook directory as AfterHours
echo F|xcopy /y VbaProject.OTM %APPDATA%\Microsoft\Outlook\AfterHours.OTM

:: change directory and save a backup of VbaProject
cd %APPDATA%\Microsoft\Outlook
@echo F|xcopy /y VbaProject.OTM VbaProject_%date:~4,2%%date:~7,2%%date:~10,4%.OTM

:: display message to user
echo .
echo .
echo .
echo =================================== 
echo Outlook must be closed in order for AfterHours to be installed.
echo =================================== 
echo .
echo .
echo .

pause

del VbaProject.OTM
ren AfterHours.OTM VbaProject.OTM

if exist AfterHours.OTM (
echo .
echo .
echo .
echo =================================== 
echo Please close Outlook and try again
echo =================================== 
) else (
echo .
echo .
echo .
echo =================================== 
echo Success!
echo =================================== 
)
pause