# vybe-test: powershell/classes_custom_methods_overloading/overload_numeric_widening_int_vs_long
class NumWidening {
    [string]Tag([int]$i) { return "Int32" }
    [string]Tag([int64]$l) { return "Int64" }
}
$nw = [NumWidening]::new()
$r1 = $nw.Tag([int]10)
$r2 = $nw.Tag([int64]10)
if ($r1 -ne "Int32" -or $r2 -ne "Int64") {
    Write-Host "FAIL: Int32 vs Int64 overload resolution failed"
    exit 1
}
Write-Host "PASS"
exit 0
