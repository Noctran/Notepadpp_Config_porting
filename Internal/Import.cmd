@Echo Off
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File "%~dp0Internal\Notepadpp_Config-Porter.ps1" -import' -Verb RunAs}"
