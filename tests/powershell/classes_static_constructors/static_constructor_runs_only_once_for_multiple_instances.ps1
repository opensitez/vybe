# vybe-test: powershell/classes_static_constructors/static_constructor_runs_only_once_for_multiple_instances
class SingleRunCheck {
    static [int]$InitCounter = 0
    static SingleRunCheck() {
        [SingleRunCheck]::InitCounter++
    }
    SingleRunCheck() {}
}
$o1 = [SingleRunCheck]::new()
$o2 = [SingleRunCheck]::new()
$o3 = [SingleRunCheck]::new()
if ([SingleRunCheck]::InitCounter -ne 1) {
    Write-Host "FAIL: Static constructor ran more than once: $([SingleRunCheck]::InitCounter)"
    exit 1
}
Write-Host "PASS"
exit 0
