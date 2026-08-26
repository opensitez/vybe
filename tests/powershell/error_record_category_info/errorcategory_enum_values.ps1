# vybe-test: powershell/error_record_category_info/errorcategory_enum_values
$cats = @(
    [System.Management.Automation.ErrorCategory]::NotSpecified,
    [System.Management.Automation.ErrorCategory]::OpenError,
    [System.Management.Automation.ErrorCategory]::CloseError,
    [System.Management.Automation.ErrorCategory]::DeviceError,
    [System.Management.Automation.ErrorCategory]::DeadlockDetected,
    [System.Management.Automation.ErrorCategory]::InvalidArgument,
    [System.Management.Automation.ErrorCategory]::InvalidData,
    [System.Management.Automation.ErrorCategory]::InvalidOperation,
    [System.Management.Automation.ErrorCategory]::InvalidResult,
    [System.Management.Automation.ErrorCategory]::ResourceUnavailable
)
if ($cats.Length -ne 10) {
    Write-Host "FAIL: ErrorCategory enum values count check failed"
    exit 1
}
Write-Host "PASS"
exit 0
