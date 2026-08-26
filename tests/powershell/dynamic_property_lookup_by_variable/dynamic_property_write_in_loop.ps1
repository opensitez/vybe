# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_write_in_loop
$obj = [pscustomobject]@{ A = 0; B = 0; C = 0 }
$props = @("A", "B", "C")
$val = 1
foreach ($p in $props) {
    $obj.$p = $val * 10
    $val++
}
if ($obj.A -ne 10 -or $obj.B -ne 20 -or $obj.C -ne 30) {
    Write-Host "FAIL: Dynamic property write in loop failed"
    exit 1
}
Write-Host "PASS"
exit 0
