# vybe-test: powershell/classes/class_base_constructor_call
class Vehicle {
    [string]$Make
    Vehicle([string]$make) { $this.Make = $make }
}
class Car : Vehicle {
    [int]$Doors
    Car([string]$make, [int]$doors) : base($make) { $this.Doors = $doors }
}
$c = [Car]::new("Toyota", 4)
if ($c.Make -ne "Toyota") { Write-Host "FAIL: Make"; exit 1 }
if ($c.Doors -ne 4)       { Write-Host "FAIL: Doors"; exit 1 }
Write-Host "PASS"
exit 0
