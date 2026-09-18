# vybe-test: powershell/split_path_cmdlet/split_path_noqualifier_windows_drive
# Split-Path -NoQualifier strips the drive or qualifier and returns the remainder of the path
$res = Split-Path "C:\Data\file.csv" -NoQualifier

if ($res -ne "\Data\file.csv") {
    Write-Host "FAIL: expected '\Data\file.csv', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
