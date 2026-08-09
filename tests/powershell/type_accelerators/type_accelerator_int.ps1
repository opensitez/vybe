# vybe-test: powershell/type_accelerators/type_accelerator_int
$i = [int]"42"
if ($i -ne 42) {
    Write-Host "FAIL: int expected 42, got $i"
    exit 1
}
if (-not ($i -is [int])) {
    Write-Host "FAIL: variable is not [int]"
    exit 1
}
Write-Host "PASS"
exit 0
