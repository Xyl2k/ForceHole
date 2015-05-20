@echo off
color 17
cls
set target=46.20.35.71/files/
set droppath=files
set start=1
set end=1000
set step=1
if not exist %droppath% (
mkdir %droppath% )
FOR /L %%G IN (%start%, %step%, %end%) DO wget -U "Mozilla/4.0 (compatible; MSIE 5.01; Windows NT 5.0)" -S -t 100 -P / "%target%%%G" -O "%droppath%/%%G"
FOR %%i IN (%droppath%\*) do if %%~zi LEQ 2 DEL %%i
echo Done.
pause