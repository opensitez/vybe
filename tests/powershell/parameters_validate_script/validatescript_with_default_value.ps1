# vybe-test: powershell/parameters_validate_script/validatescript_with_default_value
function Get-ValidatedTimeout {
    param([ValidateScript({ $_ -ge 10 -and $_ -le 60 })][int]$Timeout = 30)
    return $Timeout
}
$res = Get-ValidatedTimeout
if ($res -ne 30) {
    Write-Host "FAIL: ValidateScript default value failed"
    exit 1
}
Write-Host "PASS"
exit 0
