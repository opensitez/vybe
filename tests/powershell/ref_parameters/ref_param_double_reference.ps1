# vybe-test: powershell/ref_parameters/ref_param_double_reference
$v = "Deep"
$r1 = [ref]$v
$r2 = [ref]$r1.Value
$r2.Value = "Modified"
if ($v -ne "Modified") {
    Write-Host "FAIL: double reference mutation expected Modified, got $v"
    exit 1
}
Write-Host "PASS"
exit 0
