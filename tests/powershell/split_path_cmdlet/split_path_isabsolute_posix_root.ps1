# vybe-test: powershell/split_path_cmdlet/split_path_isabsolute_posix_root
# -IsAbsolute tests whether a path begins with a root separator or drive specification
$res = Split-Path "/var/log" -IsAbsolute

if ($res -ne $true) {
    Write-Host "FAIL: expected `$true for absolute POSIX path '/var/log', got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
