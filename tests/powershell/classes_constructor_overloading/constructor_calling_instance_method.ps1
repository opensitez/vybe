# vybe-test: powershell/classes_constructor_overloading/constructor_calling_instance_method
class Initializer {
    [int]$Total
    Initializer([int]$a, [int]$b) {
        $this.Total = $this.Calculate($a, $b)
    }
    [int]Calculate([int]$x, [int]$y) {
        return $x * $y
    }
}
$init = [Initializer]::new(6, 7)
if ($init.Total -ne 42) {
    Write-Host "FAIL: Constructor calling instance method failed, got $($init.Total)"
    exit 1
}
Write-Host "PASS"
exit 0
