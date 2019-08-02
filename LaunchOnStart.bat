@echo off

Rem ***** Change directory to Downloads and remove all files and then change directory to Onedrive - Atos *****
cd ..\..\Downloads
del *.* /f /q
cd ..\Onedrive - Atos

Rem ***** Open following tabs in Chrome *****
start Chrome
start "Remedy" "Remedy url"
start "ChaRM" "ChaRM url"
start "SAP Router" "SAPRouter url"

Rem ***** Open VPN in Internet Explorer *****
start iexplore.exe "VPN url"

Rem ***** Open Skype (Lync), Outlook, VPN Access and SAP Logon *****
start /min Lync
start /min Outlook
start VIPUIManager
start saplogon

Rem ***** Open .docm file with passwords *****
start Passwords.docm

Rem ***** Change folder to AMS_ARDAGH *****
cd AMS_ARDAGH

Rem ***** Get today's date *****
set dayToday=%Date:~8,2%
set monthToday=%Date:~5,2%
set yearToday=%Date:~0,4%
set fullDate=%yearToday%_%monthToday%_%dayToday%

Rem ***** Get yesterday's date *****
set/a dayYesterday=%dayToday%-1

Rem ***** If day is less than 10, add 0 before day number *****
if %dayYesterday% LSS 10 (
	set fullDateYesterday=%yearToday%_%monthToday%_0%dayYesterday%
) else  (
	set fullDateYesterday=%yearToday%_%monthToday%_%dayYesterday%
)

Rem ***** Set file names *****
set fileName=AMS_ARDAGH_%fullDate%.xlsm
set fileNameYesterday=AMS_ARDAGH_%fullDateYesterday%.xlsm

Rem ***** Check if file exists and if not copy file from yesterday *****
if exist %fileName% (
	start %fileName%
) else (
	copy /y %fileNameYesterday% %fileName%
	start %fileName%
)
