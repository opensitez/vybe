# vybe-test: powershell/type_accelerators/type_accelerator_array
$arr = [array](1, 2, 3)
if (-not ($arr -is [array])) {
    Write-Host "FAIL: expected [array] type"
    exit 1
}
if ($arr.Length -ne 3) {
    Write-Host "FAIL: array Length expected 3, got $($arr.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
