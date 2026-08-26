# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_with_empty_array_is_not_null
$arr = @()
# Empty array is NOT null, ?? must return empty array
$res = $arr ?? @(1, 2, 3)
if ($res.Length -ne 0) {
    Write-Host "FAIL: Empty array should not be coalesced, got length $($res.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
