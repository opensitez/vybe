# vybe-test: powershell/parameters_validate_range/validaterange_scriptblock_invocation
$sb = {
    param([ValidateRange(100, 200)][int]$Port)
    return $Port
}
$res = & $sb -Port 150
if ($res -ne 150) {
    Write-Host "FAIL: ScriptBlock ValidateRange failed"
    exit 1
}
Write-Host "PASS"
exit 0
