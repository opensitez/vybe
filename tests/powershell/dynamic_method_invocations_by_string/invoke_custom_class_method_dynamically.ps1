# vybe-test: powershell/dynamic_method_invocations_by_string/invoke_custom_class_method_dynamically
class DynamicTarget {
    [string]SayHi([string]$name) { return "Hi, $name" }
    [string]SayBye([string]$name) { return "Bye, $name" }
}
$target = [DynamicTarget]::new()
$m1 = "SayHi"
$m2 = "SayBye"
$r1 = $target.$m1("Alice")
$r2 = $target.$m2("Bob")
if ($r1 -ne "Hi, Alice" -or $r2 -ne "Bye, Bob") {
    Write-Host "FAIL: Dynamic custom class method invocation failed"
    exit 1
}
Write-Host "PASS"
exit 0
