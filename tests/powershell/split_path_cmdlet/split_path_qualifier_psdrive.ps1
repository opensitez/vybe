# vybe-test: powershell/split_path_cmdlet/split_path_qualifier_psdrive
# -Qualifier extracts the drive name from PSDrive paths (e.g. Env: or Cert:)
$res = Split-Path "Env:TEMP" -Qualifier

if ($res -ne "Env:") {
    Write-Host "FAIL: expected 'Env:', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
