# vybe-test: powershell/classes_polymorphism_and_casting/type_test_is_operator_hierarchy
class Vehicle {}
class Car : Vehicle {}
class SportsCar : Car {}
$sc = [SportsCar]::new()
if ($sc -isnot [SportsCar] -or $sc -isnot [Car] -or $sc -isnot [Vehicle] -or $sc -isnot [object]) {
    Write-Host "FAIL: -is operator hierarchy conformance failed"
    exit 1
}
Write-Host "PASS"
exit 0
