# vybe-test: powershell/error_record_category_info/categoryinfo_reason_property
$ex = [System.Exception]::new("Err")
$err = [System.Management.Automation.ErrorRecord]::new(
    $ex,
    "ErrId",
    [System.Management.Automation.ErrorCategory]::ResourceUnavailable,
    $null
)
$err.CategoryInfo.Reason = "ServiceOffline"
if ($err.CategoryInfo.Reason -ne "ServiceOffline") {
    Write-Host "FAIL: CategoryInfo Reason set failed"
    exit 1
}
Write-Host "PASS"
exit 0
