# vybe-test: powershell/classes_custom_methods_overloading/overload_with_null_argument_resolution
class NullOverload {
    [string]Process([string]$s) { return "String:$s" }
    [string]Process([object[]]$arr) { return "ArrayCount:$($arr.Length)" }
}
$no = [NullOverload]::new()
$res = $no.Process("test")
if ($res -ne "String:test") {
    Write-Host "FAIL: Overload resolution check failed"
    exit 1
}
Write-Host "PASS"
exit 0
