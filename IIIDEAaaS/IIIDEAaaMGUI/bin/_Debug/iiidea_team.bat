cd %~dp0

echo %~dp0
echo %1
echo %2
echo %3

REM %1 : FOLDER
REM %2 : CSV
REM %3 : RESULT

python iiidea_clustering.py %2
powershell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
powershell Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
powershell -File %~dp0iiidea_teaming.ps1 %1 %3

REM PAUSE 



