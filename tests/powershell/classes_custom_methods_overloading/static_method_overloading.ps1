# vybe-test: powershell/classes_custom_methods_overloading/static_method_overloading
class MathOps {
    static [int]Add([int]$a, [int]$b) { return $a + $b }
    static [double]Add([double]$a, [double]$b) { return $a + $b }
    static [string]Add([string]$a, [string]$b) { return "$a$b" }
}
$i = [MathOps]::Add(10, 20)
$d = [MathOps]::Add(1.5, 2.5)
$s = [MathOps]::Add("foo", "bar")
if ($i -ne 30 -or $d -ne 4.0 -or $s -ne "foobar") {
    Write-Host "FAIL: Static method overloading failed"
    exit 1
}
Write-Host "PASS"
exit 0
