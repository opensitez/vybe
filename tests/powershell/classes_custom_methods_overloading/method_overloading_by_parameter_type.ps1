# vybe-test: powershell/classes_custom_methods_overloading/method_overloading_by_parameter_type
class Formatter {
    [string]Format([int]$x) { return "INT:$x" }
    [string]Format([string]$x) { return "STR:$x" }
    [string]Format([bool]$x) { return "BOOL:$x" }
}
$f = [Formatter]::new()
$r1 = $f.Format(42)
$r2 = $f.Format("hello")
$r3 = $f.Format($true)
if ($r1 -ne "INT:42" -or $r2 -ne "STR:hello" -or $r3 -ne "BOOL:True") {
    Write-Host "FAIL: Method overloading by parameter type failed"
    exit 1
}
Write-Host "PASS"
exit 0
