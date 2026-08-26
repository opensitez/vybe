# vybe-test: powershell/error_record_category_info/error_record_preserves_exception_reference
$originalEx = [System.ArgumentOutOfRangeException]::new("idx", "Out of bounds")
$err = [System.Management.Automation.ErrorRecord]::new($originalEx, "RangeId", [System.Management.Automation.ErrorCategory]::LimitsExceeded, $null)
if ($err.Exception -ne $originalEx) {
    Write-Host "FAIL: ErrorRecord exception reference preservation failed"
    exit 1
}
Write-Host "PASS"
exit 0
