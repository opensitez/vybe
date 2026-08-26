# vybe-test: powershell/type_nullable_value_types/null_coalescing_with_nullable
[System.Nullable[int]]$a = $null
$res = $a ?? 10
if ($res -ne 10) {
    Write-Host "FAIL: null coalescing failed"
    exit 1
}
Write-Host "PASS"
exit 0
