# vybe-test: powershell/classes_base_method_calls/base_property_getter_access
class Vehicle {
    [int]$Speed = 50
}
class FastCar : Vehicle {
    [int]$Turbo = 30
    [int]GetTotalSpeed() {
        return ([Vehicle]$this).Speed + $this.Turbo
    }
}
$fc = [FastCar]::new()
if ($fc.GetTotalSpeed() -ne 80) {
    Write-Host "FAIL: Base property access failed"
    exit 1
}
Write-Host "PASS"
exit 0
