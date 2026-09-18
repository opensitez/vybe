# vybe-test: powershell/resolve_path_cmdlet/resolve_path_erroraction_silentlycontinue
# When -ErrorAction SilentlyContinue is set on a nonexistent path, Resolve-Path evaluates to $null without terminating
$res = Resolve-Path "missing_path_xyz_1122" -ErrorAction SilentlyContinue

if ($null -ne $res) {
    Write-Host "FAIL: expected `$null with SilentlyContinue on missing path, got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
