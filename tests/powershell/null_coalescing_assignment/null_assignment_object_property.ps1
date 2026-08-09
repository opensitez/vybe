# vybe-test: powershell/null_coalescing_assignment/null_assignment_object_property
$obj = [pscustomobject]@{ Setting = $null }
$obj.Setting ??= "DefaultSetting"
if ($obj.Setting -ne "DefaultSetting") {
    Write-Host "FAIL: object property ??= expected DefaultSetting, got $($obj.Setting)"
    exit 1
}
Write-Host "PASS"
exit 0
