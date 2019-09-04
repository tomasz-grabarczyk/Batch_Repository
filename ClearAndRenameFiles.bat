@echo off
Rem ***** Change directory to Downloads and remove all files with extentions listed in line *****
cd ..\..\Downloads
del *.msg *.eml *.pdf *.png *.jpg /f /q

Rem ***** Change name of the file to be visible for Excel Report *****
if exist *UpdateReport.xls Ren *UpdateReport.xls ArdaghDailyUpdateReport.xls

Rem ***** If there are two export.csv files change names of the files accordingly for Excel Report *****
if exist "export (1).csv" (
	Ren export.csv cd.csv
	Ren "export (1).csv" rfc.csv
)

if exist *DevCounterTomGra.xls Ren *DevCounterTomGra.xls DevCounterTomGra.xls
