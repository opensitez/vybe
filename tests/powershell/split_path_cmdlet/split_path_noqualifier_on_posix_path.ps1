# vybe-test: powershell/split_path_cmdlet/split_path_noqualifier_on_posix_path
# On paths without a drive qualifier, -NoQualifier returns the path unchanged
$res = Split-Path "/usr/local/bin" -NoQualifier

if ($res -ne "/usr/local/bin") {
    Write-Host "FAIL: expected '/usr/local/bin', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
