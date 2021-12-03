@echo off
:: change directory and save a backup of VbaProject
cd %APPDATA%\Microsoft\Outlook
@echo F|xcopy /y VbaProject.OTM VbaProject_%date:~4,2%%date:~7,2%%date:~10,4%.OTM

:: display message to user
echo .
echo .
echo .
echo =================================== 
echo Outlook must be closed in order for AfterHours to be removed.
echo =================================== 
echo .
echo .
echo .

pause

del VbaProject.OTM

if exist VbaProject.OTM (
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