REM @ECHO OFF
cls
setlocal enableextensions
set nameDir=%DATE:/=_%
mkdir %nameDir%
xcopy  .\mongorestore-dump-to-local.cmd .\%nameDir% /K /D /H /Y
mongodump --uri="mongodb://192.168.1.11:27017" --db="tiqn" --out %nameDir%
 
