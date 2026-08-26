# vybe-test: powershell/parameters_validate_pattern/validatepattern_scriptblock_parameter
$sb = {
    param([ValidatePattern('^\d+$')][string]$NumStr)
    return [int]$NumStr * 2
}
$res = & $sb -NumStr "25"
if ($res -ne 50) {
    Write-Host "FAIL: Scriptblock ValidatePattern failed"
    exit 1
}
Write-Host "PASS"
exit 0
