REM Set execution policy to unrestricted or other behaviors set by args

SET policy="UnRestricted"
IF NOT [%1]==[] SET policy=%1

PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList 'Set-ExecutionPolicy %policy% -Force' -Verb RunAs}"