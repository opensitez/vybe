# vybe-test: powershell/parameters_validate_pattern/validatepattern_alphanumeric_with_underscores
function Set-Identifier {
    param([ValidatePattern('^[a-zA-Z_]\w*$')][string]$Id)
    return $Id
}
$res = Set-Identifier -Id "_myVar1"
if ($res -ne "_myVar1") {
    Write-Host "FAIL: ValidatePattern identifier failed"
    exit 1
}
Write-Host "PASS"
exit 0
