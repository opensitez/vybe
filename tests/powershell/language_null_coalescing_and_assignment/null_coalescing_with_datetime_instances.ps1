# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_with_datetime_instances
$dt = [datetime]::UtcNow
$nullDt = $null
$res = $nullDt ?? $dt
if ($res -ne $dt) {
    Write-Host "FAIL: ?? with DateTime instances failed"
    exit 1
}
Write-Host "PASS"
exit 0
