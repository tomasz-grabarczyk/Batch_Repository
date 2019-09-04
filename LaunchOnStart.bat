@echo off

Rem ***** Change directory to Downloads and remove all files and then change directory to Onedrive - Atos *****
cd ..\..\Downloads
del *.* /f /q
cd ..\Onedrive - Atos

Rem ***** Open following tabs in Chrome *****
start Chrome
start "Remedy" "https://cgorion.onbmc.com/"
start "ChaRM" "http://ardasmci.ardaghgroup.com:8080/sap/bc/bsp/sap/crm_ui_start/default.htm?sap-client=060&sap-language=EN"
start "SAP Router" "http://g11as1.informatik.tu-muenchen.de:8011/remotelogin(bD1lbiZjPTEwMQ==)/dremlogin.htm"

Rem ***** Open VPN in Internet Explorer *****
start iexplore.exe "https://vpneu.ardaghgroup.com/my.policy"

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
if not exist %fileName% (
	copy /y %fileNameYesterday% %fileName%
)

Rem ***** Open folder with Excel files *****
explorer C:\Users\%USERNAME%\OneDrive - Atos\AMS_ARDAGH
