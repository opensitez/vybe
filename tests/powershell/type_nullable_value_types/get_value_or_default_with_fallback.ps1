# vybe-test: powershell/type_nullable_value_types/get_value_or_default_with_fallback
[System.Nullable[int]]$x = $null
if ($x.GetValueOrDefault(99) -ne 99) {
    Write-Host "FAIL: GetValueOrDefault(99) should return 99"
    exit 1
}
Write-Host "PASS"
exit 0
