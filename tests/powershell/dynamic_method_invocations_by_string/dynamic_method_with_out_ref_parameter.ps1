# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_with_out_ref_parameter
$m = "TryParse"
$outVal = 0
$ok = [int]::$m("999", [ref]$outVal)
if (-not $ok -or $outVal -ne 999) {
    Write-Host "FAIL: Dynamic static method with ref parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
