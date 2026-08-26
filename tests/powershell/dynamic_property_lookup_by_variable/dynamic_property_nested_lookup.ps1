# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_nested_lookup
$outer = "Parent"
$inner = "ChildName"
$tree = [pscustomobject]@{
    Parent = [pscustomobject]@{ ChildName = "Baby" }
}
$res = $tree.$outer.$inner
if ($res -ne "Baby") {
    Write-Host "FAIL: Nested dynamic property lookup failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
