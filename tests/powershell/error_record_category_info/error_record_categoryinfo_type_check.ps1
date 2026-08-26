# vybe-test: powershell/error_record_category_info/error_record_categoryinfo_type_check
$ex = [System.Exception]::new("Err")
$err = [System.Management.Automation.ErrorRecord]::new($ex, "Id", [System.Management.Automation.ErrorCategory]::NotSpecified, $null)
if ($err.CategoryInfo.GetType().Name -ne "ErrorCategoryInfo") {
    Write-Host "FAIL: CategoryInfo type expected ErrorCategoryInfo, got $($err.CategoryInfo.GetType().Name)"
    exit 1
}
Write-Host "PASS"
exit 0
