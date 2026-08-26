# vybe-test: powershell/error_record_category_info/error_record_error_details_message
$ex = [System.Exception]::new("BaseEx")
$err = [System.Management.Automation.ErrorRecord]::new($ex, "Id", [System.Management.Automation.ErrorCategory]::NotSpecified, $null)
$err.ErrorDetails = [System.Management.Automation.ErrorDetails]::new("Detailed custom message")
if ($err.ErrorDetails.Message -ne "Detailed custom message") {
    Write-Host "FAIL: ErrorDetails custom message check failed"
    exit 1
}
Write-Host "PASS"
exit 0
