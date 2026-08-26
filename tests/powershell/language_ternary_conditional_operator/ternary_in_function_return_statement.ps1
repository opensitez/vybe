# vybe-test: powershell/language_ternary_conditional_operator/ternary_in_function_return_statement
function Get-ModeName([bool]$isDebug) {
    return $isDebug ? "DebugMode" : "ReleaseMode"
}
$r1 = Get-ModeName $true
$r2 = Get-ModeName $false
if ($r1 -ne "DebugMode" -or $r2 -ne "ReleaseMode") {
    Write-Host "FAIL: Ternary in function return failed"
    exit 1
}
Write-Host "PASS"
exit 0
