@echo off
rem MultiReplace build script (requires .NET 8 SDK or later)
rem Output: dist\MultiReplace.exe  (runs with .NET Desktop Runtime 8+)
cd /d "%~dp0"
dotnet publish MultiReplace.csproj -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -p:DebugType=none -o dist
if errorlevel 1 (
  echo.
  echo BUILD FAILED
  pause
  exit /b 1
)
echo.
echo Done: %~dp0dist\MultiReplace.exe
pause
