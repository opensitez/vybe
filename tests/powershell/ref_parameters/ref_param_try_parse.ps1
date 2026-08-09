# vybe-test: powershell/ref_parameters/ref_param_try_parse
$outVal = 0
$success = [int]::TryParse("123", [ref]$outVal)
if (-not $success -or $outVal -ne 123) {
    Write-Host "FAIL: [int]::TryParse expected true and 123, got success=$success outVal=$outVal"
    exit 1
}
Write-Host "PASS"
exit 0
