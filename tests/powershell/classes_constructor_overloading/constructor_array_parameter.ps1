# vybe-test: powershell/classes_constructor_overloading/constructor_array_parameter
class DataSet {
    [int[]]$Numbers
    DataSet([int[]]$nums) { $this.Numbers = $nums }
}
$ds = [DataSet]::new(@(1, 2, 3))
if ($ds.Numbers.Length -ne 3 -or $ds.Numbers[1] -ne 2) {
    Write-Host "FAIL: Array parameter constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
