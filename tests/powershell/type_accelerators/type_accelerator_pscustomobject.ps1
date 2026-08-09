# vybe-test: powershell/type_accelerators/type_accelerator_pscustomobject
$obj = [pscustomobject]@{ Name = "Vybe"; Speed = 100 }
if ($obj.Name -ne "Vybe") {
    Write-Host "FAIL: Name expected Vybe, got $($obj.Name)"
    exit 1
}
if ($obj.Speed -ne 100) {
    Write-Host "FAIL: Speed expected 100, got $($obj.Speed)"
    exit 1
}
Write-Host "PASS"
exit 0
