# vybe-test: powershell/ref_parameters/ref_param_value_property
$v = 100
$refObj = [ref]$v
if ($refObj.Value -ne 100) {
    Write-Host "FAIL: direct [ref] object Value property expected 100, got $($refObj.Value)"
    exit 1
}
$refObj.Value = 200
if ($v -ne 200) {
    Write-Host "FAIL: mutation via .Value property expected 200, got $v"
    exit 1
}
Write-Host "PASS"
exit 0
