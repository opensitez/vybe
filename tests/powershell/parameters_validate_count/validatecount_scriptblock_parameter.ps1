# vybe-test: powershell/parameters_validate_count/validatecount_scriptblock_parameter
$sb = {
    param([ValidateCount(1, 2)][string[]]$Items)
    return $Items -join ","
}
$res = & $sb -Items "A", "B"
if ($res -ne "A,B") {
    Write-Host "FAIL: ScriptBlock ValidateCount failed"
    exit 1
}
Write-Host "PASS"
exit 0
