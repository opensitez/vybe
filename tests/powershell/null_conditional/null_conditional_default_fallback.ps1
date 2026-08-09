# vybe-test: powershell/null_conditional/null_conditional_default_fallback
$obj = $null
$res = (${obj}?.Value) ?? "Default"
if ($res -ne "Default") {
    Write-Host "FAIL: null-conditional with coalescing fallback expected 'Default', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
