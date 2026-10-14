@echo off
title PDF Araci kaldir
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Process -Name \"PDF Araci\" -EA 0 | Stop-Process -Force -EA 0; Remove-Item \"$env:LOCALAPPDATA\PdfAraci\",\"$env:APPDATA\PdfAraci\" -Recurse -Force -EA 0; Get-ChildItem ([Environment]::GetFolderPath(\"Desktop\")),([Environment]::GetFolderPath(\"Programs\")) -Filter \"PDF Arac*.lnk\" -EA 0 | Remove-Item -Force -EA 0; Write-Host \"PDF Araci kaldirildi.\""
pause
