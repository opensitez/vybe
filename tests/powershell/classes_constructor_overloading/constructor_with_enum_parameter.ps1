# vybe-test: powershell/classes_constructor_overloading/constructor_with_enum_parameter
enum Level { Low; High }
class PriorityItem {
    [Level]$Pri
    PriorityItem([Level]$p) { $this.Pri = $p }
}
$pi = [PriorityItem]::new([Level]::High)
if ($pi.Pri -ne [Level]::High) {
    Write-Host "FAIL: Constructor with enum parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
