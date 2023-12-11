# Import-Module Microsoft.PowerShell.Utility
$pythonScriptsPath = Join-Path $env:APPDATA "Python\Python310\Scripts"
$env:Path += ";$pythonScriptsPath"

$params = "--onefile  --hidden-import openpyxl"
$command = "pyinstaller .\searchFoldersAnalyzer.py $params --name searchFoldersAnalyzer.exe"

Invoke-Expression $command