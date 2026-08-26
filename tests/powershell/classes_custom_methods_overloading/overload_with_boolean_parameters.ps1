# vybe-test: powershell/classes_custom_methods_overloading/overload_with_boolean_parameters
class BoolOverloadTarget {
    [string]Execute([bool]$flag) { return "Bool:$flag" }
    [string]Execute([int]$n) { return "Int:$n" }
}
$t = [BoolOverloadTarget]::new()
$r1 = $t.Execute($true)
$r2 = $t.Execute(42)
if ($r1 -ne "Bool:True" -or $r2 -ne "Int:42") {
    Write-Host "FAIL: Boolean method overload failed"
    exit 1
}
Write-Host "PASS"
exit 0
