# vybe-test: powershell/type_nullable_value_types/nullable_in_class_field
class Container {
    [System.Nullable[int]]$Count
}
$c1 = [Container]::new()
$c1.Count = 7
$c2 = [Container]::new()
$c2.Count = $null
if ($c1.Count -ne 7 -or $null -ne $c2.Count) {
    Write-Host "FAIL: Class nullable field initialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
