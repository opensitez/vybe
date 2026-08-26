# vybe-test: powershell/error_record_category_info/create_custom_error_record_with_category
$ex = [System.InvalidOperationException]::new("Operation failed")
$err = [System.Management.Automation.ErrorRecord]::new(
    $ex,
    "OperationFailedErrorId",
    [System.Management.Automation.ErrorCategory]::InvalidOperation,
    "TargetObjectData"
)
if ($err.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::InvalidOperation) {
    Write-Host "FAIL: ErrorRecord category check failed"
    exit 1
}
Write-Host "PASS"
exit 0
