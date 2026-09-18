# vybe-test: powershell/split_path_cmdlet/split_path_isabsolute_windows_drive
# Windows drive paths like C:\Windows are parsed as absolute
$res = Split-Path "C:\Windows" -IsAbsolute

if ($res -ne $true) {
    Write-Host "FAIL: expected `$true for Windows drive path 'C:\Windows', got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
