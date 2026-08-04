# vybe-test: powershell/classes/class_property_static
class Counter {
    static [int]$Count = 0
    
    Counter() {
        [Counter]::Count++
    }
}

$c1 = [Counter]::new()
$c2 = [Counter]::new()
$c3 = [Counter]::new()

if ([Counter]::Count -ne 3) {
    Write-Host "FAIL: expected static Count = 3, got $([Counter]::Count)"
    exit 1
}
Write-Host "PASS"
exit 0
