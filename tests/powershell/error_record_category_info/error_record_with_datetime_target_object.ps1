# vybe-test: powershell/error_record_category_info/error_record_with_datetime_target_object
$dt = [datetime]::UtcNow
$ex = [System.Exception]::new("DateErr")
$err = [System.Management.Automation.ErrorRecord]::new($ex, "DId", [System.Management.Automation.ErrorCategory]::LimitsExceeded, $dt)
if ($err.TargetObject -ne $dt) {
    Write-Host "FAIL: ErrorRecord with DateTime target failed"
    exit 1
}
Write-Host "PASS"
exit 0
