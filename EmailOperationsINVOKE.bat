@echo off
cd /D "%~dp0"
goto check_Permissions

:check_Permissions
 echo Administrative permissions required. Detecting permissions...

 net session >nul 2>&1
 if %errorLevel% == 0 (
 echo Success: Administrative permissions confirmed.

 ) else (
 echo Failure: Current permissions inadequate.
 echo Please restart with administrative permissions
 pause
 exit
 )

:SEARCH

SET /P searchterm=Please enter an e-mail address or Subject (prepended and appended wildcards are assumed for SUBJECT searches, emails are assumed exact)

SET M=Undefined
SET mode=search



:MENU
cls
ECHO Configured to search by: %M%
ECHO Mode is set to: %mode%
ECHO Search Term is set to: %searchterm%
IF %mode%==murder ECHO !!!! WARNING YOU ARE IN MURDER MODE--THIS SHOULD MAKE YOU AFRAID !!!!
ECHO.
ECHO ...............................................
ECHO PRESS 1, 2, 3, 5, 6, 7 or 8
ECHO ...............................................
ECHO.
ECHO 1 - Sender
ECHO 2 - Recipient -- Search Only (will ignore "murder mode")
ECHO 3 - Subject
ECHO.
ECHO 4 - Switch to Murder Mode
ECHO 5 - Switch to Search Mode
ECHO.
ECHO 6 - Change Search Term
ECHO.
ECHO 7 - View Existing Audit Search Status
ECHO 8 - View Results of an Audit Search
ECHO.
SET /P M=Type 1, 2, 3, 4, 5, 6, 7, or 8 then press ENTER:


IF %M%==1 (
   IF %mode%==murder (
     ECHO You are in SENDER and MURDER mode
     GOTO sendermurder
   )
)

IF %M%==3 (
   IF %mode%==murder (
     ECHO You are in SUBJECTand MURDER mode
     GOTO subjectmurder
   )
)


IF %M%==1 set M=sender && GOTO SENDER
IF %M%==2 set M=recipient && GOTO RECIPIENT
IF %M%==3 set M=subject && GOTO SUBJECT
IF %M%==4 set mode=murder && GOTO MENU
IF %M%==5 set mode=search && GOTO MENU
IF %M%==6 GOTO SEARCH
IF %M%==7 GOTO CHECK
IF %M%==8 GOTO REVIEW

echo You picked: %M%
echo Please select 1, 2 3, 4, 5, 6, 7, or 8. 
echo Yes. I do Oxford commas. And you suck at picking things. 
echo Returning to menu
pause

GOTO MENU

:SENDER

cls
ECHO Configured to search by: %M%
ECHO Mode is set to: %mode%
ECHO Search Term is set to: %searchterm%
ECHO.
ECHO ...............................................
ECHO PRESS 1, or 2
ECHO To set Search or Delete mode
ECHO ...............................................
ECHO.
ECHO 1 - Basic (~3 days of data, runs FAST)
ECHO 2 - Full Audit Report
ECHO.
SET /P S=Type 1, or 2 then press ENTER:


IF %S%==1 GOTO senderbasic 
IF %S%==2 GOTO senderaudit

ECHO You picked %S%
ECHO WTF are you doing?
echo Please pick 1, or 2
pause
GOTO SENDER

:SENDERBASIC

cls
ECHO Generating a basic on-screen report...
powershell -Command "& {set-variable -name x -value $env:m;set-variable -name y -value $env:searchterm;echo 'Command type is: '; echo $x;echo 'Search Term is: ' $y; if (!(Get-Module "ExchangeOnlineManagement")) {echo 'module is not loaded';echo 'Installing and/or loading module...'; Install-Module -Name ExchangeOnlineManagement;import-module -name ExchangeOnlineManagement; Connect-ExchangeOnline; get-messagetrace -senderaddress $y;echo 'Connecting to security and compliance center...';Connect-IPPSSession} else {echo 'Module Exists and is loaded';Connect-ExchangeOnline;get-messagetrace -senderaddress $y}}"

pause
SET M=Undefined
SET mode=search
GOTO MENU

:SENDERAUDIT

cls
powershell -Command "& {set-variable -name x -value $env:m; echo 'Command type is: '; echo $x; set-variable -name y -value $env:searchterm; if (!(Get-Module "ExchangeOnlineManagement")) {echo 'module is not loaded';echo 'Installing and/or loading Exchange Online Management module...'; Install-Module -Name ExchangeOnlineManagement;import-module -name ExchangeOnlineManagement; Connect-ExchangeOnline;Connect-IPPSSession;[string]$date=get-date; New-ComplianceSearch -Name "$date" -ExchangeLocation all -ContentMatchQuery from:$y;start-compliancesearch $date} else {echo 'Module Exists and is loaded';Connect-ExchangeOnline;Connect-IPPSSession;[string]$date=get-date; New-ComplianceSearch -Name "$date" -ExchangeLocation all -ContentMatchQuery ‑ContentMatchQuery from:$y;start-compliancesearch $date}}"

pause
SET M=Undefined
SET mode=search
GOTO MENU

:RECIPIENT
cls
ECHO This module is limited to about three days of data in the interest of immediete results. This is also read only. It uses Message Trace instead of ComplianceSearch.  
ECHO .
pause
powershell -Command "& {set-variable -name x -value $env:m;set-variable -name y -value $env:searchterm;echo 'Command type is: '; echo $x;echo 'Search Term is: ' $y; if (!(Get-Module "ExchangeOnlineManagement")) {echo 'module is not loaded';echo 'Installing and/or loading module...'; Install-Module -Name ExchangeOnlineManagement;import-module -name ExchangeOnlineManagement; Connect-ExchangeOnline; get-messagetrace -recipientaddress $y} else {echo 'Module Exists and is loaded';Connect-ExchangeOnline;get-messagetrace -recipientaddress $y}}"
SET M=Undefined
SET mode=search
pause

GOTO MENU

:SUBJECT
cls
powershell -Command "& {set-variable -name x -value $env:m; echo 'Command type is: '; echo $x; set-variable -name y -value $env:searchterm; if (!(Get-Module "ExchangeOnlineManagement")) {echo 'module is not loaded';echo 'Installing and/or loading Exchange Online Management module...'; Install-Module -Name ExchangeOnlineManagement;import-module -name ExchangeOnlineManagement; Connect-ExchangeOnline;Connect-IPPSSession;[string]$date=get-date; New-ComplianceSearch -Name "$date" -ExchangeLocation all -ContentMatchQuery subject:`'$y`';start-compliancesearch "$date"} else {echo 'Module Exists and is loaded';Connect-ExchangeOnline;Connect-IPPSSession;[string]$date=get-date; New-ComplianceSearch -Name "$date" -ExchangeLocation all -ContentMatchQuery subject:`'$y`';start-compliancesearch "$date"}}"
SET M=Undefined
SET mode=search
pause
GOTO MENU


:CHECK
cls
powershell -Command "& {set-variable -name x -value $env:m; if (!(Get-Module "ExchangeOnlineManagement")) {echo 'module is not loaded';echo 'Installing and/or loading Exchange Online Management module...'; Install-Module -Name ExchangeOnlineManagement;import-module -name ExchangeOnlineManagement; Connect-ExchangeOnline;Connect-IPPSSession;get-ComplianceSearch} else {echo 'Module Exists and is loaded';Connect-ExchangeOnline;Connect-IPPSSession;get-ComplianceSearch}}"
SET M=Undefined
SET mode=search
pause
GOTO MENU

:SENDERMURDER
ECHO Initializing SENDERMURDER mode...
ECHO If you don't plan on deleting data CLOSE THIS WINDOW
pause
powershell -Command "& {set-variable -name x -value $env:m; if (!(Get-Module "ExchangeOnlineManagement")) {echo 'module is not loaded';echo 'Installing and/or loading Exchange Online Management module...'; Install-Module -Name ExchangeOnlineManagement;import-module -name ExchangeOnlineManagement; Connect-ExchangeOnline;Connect-IPPSSession;Get-Compliancesearch;start-sleep -s 60;echo "Working";$reportout = Read-Host -Prompt 'Input the name of the report exactly as it apears (eg. 07/19/2021 13:34:14)';ECHO 'This is where the murder happens, but this is a placeholder. You would have deleted'; echo $reportout} else {echo 'Module Exists and is loaded';Connect-ExchangeOnline;Connect-IPPSSession;Get-Compliancesearch;start-sleep -s 60;echo "Working";$reportout = Read-Host -Prompt 'Input the name of the report exactly as it apears (eg. 07/19/2021 13:34:14)';ECHO 'This is where the murder happens, but this is a placeholder. You would have deleted';echo $reportout}}"
pause
SET M=Undefined
SET mode=search
GOTO MENU

:SUBJECTMURDER
ECHO Initializing SUBJECTMURDER mode...
ECHO If you don't plan on deleting data CLOSE THIS WINDOW
pause
powershell -Command "& {set-variable -name x -value $env:m; if (!(Get-Module "ExchangeOnlineManagement")) {echo 'module is not loaded';echo 'Installing and/or loading Exchange Online Management module...'; Install-Module -Name ExchangeOnlineManagement;import-module -name ExchangeOnlineManagement; Connect-ExchangeOnline;Connect-IPPSSession;Get-Compliancesearch;start-sleep -s 60;echo "Working";$reportout = Read-Host -Prompt 'Input the name of the report exactly as it apears (eg. 07/19/2021 13:34:14)';ECHO 'This is where the murder happens, but this is a placeholder, you would have deleted'; echo $reportout7/19} else {echo 'Module Exists and is loaded';Connect-ExchangeOnline;Connect-IPPSSession;Get-Compliancesearch;start-sleep -s 60;echo "Working";$reportout = Read-Host -Prompt 'Input the name of the report exactly as it apears (eg. 07/19/2021 13:34:14)';ECHO 'This is where the murder happens, but this is a placeholder, you would have deleted'; echo $reportout}}"

pause
SET M=Undefined
SET mode=search
GOTO MENU



:REVIEW
cls
powershell -Command "& {set-variable -name x -value $env:m; if (!(Get-Module "ExchangeOnlineManagement")) {echo 'module is not loaded';echo 'Installing and/or loading Exchange Online Management module...'; Install-Module -Name ExchangeOnlineManagement;import-module -name ExchangeOnlineManagement; Connect-ExchangeOnline;Connect-IPPSSession;get-ComplianceSearch;start-sleep -s 60;echo "Working...";$reportout = Read-Host -Prompt 'Input the name of the report exactly as it apears (eg. 07/19/2021 13:34:14)';Get-Compliancesearch -identity $reportout | select results} else {echo 'Module Exists and is loaded';Connect-ExchangeOnline;Connect-IPPSSession;get-ComplianceSearch;start-sleep -s 60;echo "Working";$reportout = Read-Host -Prompt 'Input the name of the report exactly as it apears (eg. 07/19/2021 13:34:14)';Get-Compliancesearch -identity $reportout | fl; Get-Compliancesearch -identity $reportout | fl; get-compliancesearch -identity $reportout | fl -property items}}"

pause
SET M=Undefined
SET mode=search
GOTO MENU

:ENDING

ECHO Done.
pause

REM PowerShell.exe -ExecutionPolicy Bypass -File .\file.ps1
REM New-ComplianceSearchAction -SearchName $reportout -Purge -PurgeType HardDelete
