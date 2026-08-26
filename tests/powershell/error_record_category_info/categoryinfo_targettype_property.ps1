# vybe-test: powershell/error_record_category_info/categoryinfo_targettype_property
$target = [pscustomobject]@{ Key = 123 }
$ex = [System.Exception]::new("Err")
$err = [System.Management.Automation.ErrorRecord]::new(
    $ex,
    "ErrId",
    [System.Management.Automation.ErrorCategory]::ObjectNotFound,
    $target
)
if ($err.CategoryInfo.TargetType -ne "PSCustomObject" -and $err.CategoryInfo.TargetType -ne "PSObject") {
    Write-Host "FAIL: CategoryInfo TargetType check failed, got '$($err.CategoryInfo.TargetType)'"
    exit 1
}
Write-Host "PASS"
exit 0
