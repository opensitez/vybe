# vybe-test: powershell/type_accelerators/type_accelerator_psvariable
$v = [psvariable]::new("myVar", 100)
if ($v.Name -ne "myVar") {
    Write-Host "FAIL: psvariable Name expected myVar, got $($v.Name)"
    exit 1
}
if ($v.Value -ne 100) {
    Write-Host "FAIL: psvariable Value expected 100, got $($v.Value)"
    exit 1
}
Write-Host "PASS"
exit 0
