# vybe-test: powershell/classes_custom_methods_overloading/method_overloading_array_vs_scalar
class ArrayOverload {
    [int]Sum([int]$x) { return $x }
    [int]Sum([int[]]$arr) {
        $total = 0
        foreach ($i in $arr) { $total += $i }
        return $total
    }
}
$ao = [ArrayOverload]::new()
$s1 = $ao.Sum(10)
$s2 = $ao.Sum(@(1, 2, 3, 4))
if ($s1 -ne 10 -or $s2 -ne 10) {
    Write-Host "FAIL: Array vs scalar method overload failed"
    exit 1
}
Write-Host "PASS"
exit 0
