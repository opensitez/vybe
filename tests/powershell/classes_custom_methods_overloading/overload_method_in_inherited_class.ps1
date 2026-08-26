# vybe-test: powershell/classes_custom_methods_overloading/overload_method_in_inherited_class
class BaseAdder {
    [int]Add([int]$a, [int]$b) { return $a + $b }
}
class DerivedAdder : BaseAdder {
    [int]Add([int]$a, [int]$b, [int]$c) { return $a + $b + $c }
}
$da = [DerivedAdder]::new()
if ($da.Add(1, 2) -ne 3 -or $da.Add(1, 2, 3) -ne 6) {
    Write-Host "FAIL: Overloading across inheritance boundary failed"
    exit 1
}
Write-Host "PASS"
exit 0
