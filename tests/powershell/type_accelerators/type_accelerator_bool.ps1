# vybe-test: powershell/type_accelerators/type_accelerator_bool
$b = [bool]"true"
if ($b -ne $true) {
    Write-Host "FAIL: bool expected true, got $b"
    exit 1
}
Write-Host "PASS"
exit 0
