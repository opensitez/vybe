# vybe-test: powershell/classes/class_static_member
class Counter {
    static [int]$Count = 0
    static [void]Increment() { [Counter]::Count++ }
    static [int]GetCount() { return [Counter]::Count }
}
[Counter]::Increment()
[Counter]::Increment()
[Counter]::Increment()
$val = [Counter]::GetCount()
if ($val -ne 3) {
    Write-Host "FAIL: expected 3, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
