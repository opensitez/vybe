# vybe-test: powershell/parameters_validate_length/validatelength_scriptblock_parameter
$sb = {
    param([ValidateLength(2, 6)][string]$Tag)
    return $Tag.ToUpper()
}
$res = & $sb -Tag "prod"
if ($res -ne "PROD") {
    Write-Host "FAIL: ScriptBlock ValidateLength failed"
    exit 1
}
Write-Host "PASS"
exit 0
