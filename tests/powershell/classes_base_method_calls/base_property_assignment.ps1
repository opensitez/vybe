# vybe-test: powershell/classes_base_method_calls/base_property_assignment
class BaseData {
    [string]$Title
}
class SubData : BaseData {
    [void]SetBaseTitle([string]$t) {
        ([BaseData]$this).Title = $t
    }
}
$sd = [SubData]::new()
$sd.SetBaseTitle("NewTitle")
if ($sd.Title -ne "NewTitle") {
    Write-Host "FAIL: Base property assignment failed"
    exit 1
}
Write-Host "PASS"
exit 0
