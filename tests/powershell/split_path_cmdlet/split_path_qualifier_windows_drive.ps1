# vybe-test: powershell/split_path_cmdlet/split_path_qualifier_windows_drive
# Split-Path -Qualifier returns the drive qualifier (e.g. 'C:')
$res = Split-Path "C:\Data\file.csv" -Qualifier

if ($res -ne "C:") {
    Write-Host "FAIL: expected 'C:', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
