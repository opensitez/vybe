# vybe-test: powershell/error_record_category_info/categoryinfo_activity_property
$ex = [System.Exception]::new("Err")
$err = [System.Management.Automation.ErrorRecord]::new(
    $ex,
    "ErrId",
    [System.Management.Automation.ErrorCategory]::PermissionDenied,
    $null
)
$err.CategoryInfo.Activity = "WriteToFile"
if ($err.CategoryInfo.Activity -ne "WriteToFile") {
    Write-Host "FAIL: CategoryInfo Activity set failed"
    exit 1
}
Write-Host "PASS"
exit 0
