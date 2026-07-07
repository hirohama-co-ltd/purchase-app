@echo off
setlocal
cd /d %TEMP%
pushd "%~dp0"
if errorlevel 1 (
  echo pushd failed: %~dp0
  exit /b 1
)
echo PWD: %CD%
echo.
echo === clasp push ===
call clasp.cmd push --force
if errorlevel 1 (
  popd
  exit /b 1
)
echo.
echo === clasp deploy ===
call clasp.cmd deploy -i AKfycbz6_wkmOhXz2JsEhjgfeDyaDoHNEZNFUrPoIIpBfFDmjKkbTT9IKeR0pwsIHSZ5dRu7kw -d "DOMAIN公開"
set ERR=%ERRORLEVEL%
popd
exit /b %ERR%
