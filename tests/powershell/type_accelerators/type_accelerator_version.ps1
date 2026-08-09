# vybe-test: powershell/type_accelerators/type_accelerator_version
$v = [version]"1.2.3.4"
if ($v.Major -ne 1) {
    Write-Host "FAIL: Major expected 1, got $($v.Major)"
    exit 1
}
if ($v.Minor -ne 2) {
    Write-Host "FAIL: Minor expected 2, got $($v.Minor)"
    exit 1
}
if ($v.Build -ne 3) {
    Write-Host "FAIL: Build expected 3, got $($v.Build)"
    exit 1
}
if ($v.Revision -ne 4) {
    Write-Host "FAIL: Revision expected 4, got $($v.Revision)"
    exit 1
}
Write-Host "PASS"
exit 0
