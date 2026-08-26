# vybe-test: powershell/classes_polymorphism_and_casting/polymorphic_property_getter_simulation
class Metric {
    [int]GetScore() { return 10 }
}
class BoostedMetric : Metric {
    [int]GetScore() { return 50 }
}
[Metric]$m = [BoostedMetric]::new()
if ($m.GetScore() -ne 50) {
    Write-Host "FAIL: Polymorphic score getter failed"
    exit 1
}
Write-Host "PASS"
exit 0
