# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_dispatch_based_on_runtime_condition
function Format-StringDynamic([string]$inputStr, [bool]$upper) {
    $action = if ($upper) { "ToUpper" } else { "ToLower" }
    return $inputStr.$action()
}
$r1 = Format-StringDynamic "Test" $true
$r2 = Format-StringDynamic "Test" $false
if ($r1 -ne "TEST" -or $r2 -ne "test") {
    Write-Host "FAIL: Conditional dynamic method dispatch failed"
    exit 1
}
Write-Host "PASS"
exit 0
