# vybe-test: powershell/convert_path_cmdlet/convert_path_erroraction_silentlycontinue
# With -ErrorAction SilentlyContinue, a non-existent path suppresses the error and evaluates to $null
$res = Convert-Path "completely_missing_path_xyz_8877" -ErrorAction SilentlyContinue

if ($null -ne $res) {
    Write-Host "FAIL: expected `$null with SilentlyContinue on missing path, got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
