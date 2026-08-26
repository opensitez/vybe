# vybe-test: powershell/type_nullable_value_types/nullable_in_array_collection
$arr = [System.Array]::CreateInstance([type]"System.Nullable[int]", 3)
$arr.SetValue([System.Nullable[int]]1, 0)
$arr.SetValue($null, 1)
$arr.SetValue([System.Nullable[int]]3, 2)
if ($arr.GetValue(0) -ne 1 -or $arr.GetValue(1) -ne $null -or $arr.GetValue(2) -ne 3) {
    Write-Host "FAIL: Nullable array collection check failed"
    exit 1
}
Write-Host "PASS"
exit 0
