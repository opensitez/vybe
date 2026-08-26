# vybe-test: powershell/parameters_validate_script/validatescript_multiple_parameters
function Multi-ScriptParam {
    param(
        [ValidateScript({ $_ -gt 0 })][int]$X,
        [ValidateScript({ $_ -lt 0 })][int]$Y
    )
    return "$X,$Y"
}
$res = Multi-ScriptParam -X 5 -Y -5
if ($res -ne "5,-5") {
    Write-Host "FAIL: Multiple ValidateScript parameters failed"
    exit 1
}
Write-Host "PASS"
exit 0
