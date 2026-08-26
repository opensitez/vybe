# vybe-test: powershell/error_record_category_info/categoryinfo_targetname_property
$ex = [System.ArgumentException]::new("Bad arg")
$err = [System.Management.Automation.ErrorRecord]::new(
    $ex,
    "BadArgId",
    [System.Management.Automation.ErrorCategory]::InvalidArgument,
    "MyConfigPath"
)
if ($err.CategoryInfo.TargetName -ne "MyConfigPath") {
    Write-Host "FAIL: CategoryInfo TargetName check failed, got '$($err.CategoryInfo.TargetName)'"
    exit 1
}
Write-Host "PASS"
exit 0
