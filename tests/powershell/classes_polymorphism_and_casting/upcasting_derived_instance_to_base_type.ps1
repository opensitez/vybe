# vybe-test: powershell/classes_polymorphism_and_casting/upcasting_derived_instance_to_base_type
class Animal {
    [string]Speak() { return "Sound" }
}
class Cat : Animal {
    [string]Speak() { return "Meow" }
}
$c = [Cat]::new()
[Animal]$a = $c
if ($a.Speak() -ne "Meow") {
    Write-Host "FAIL: Virtual method dispatch via upcasted variable failed, got '$($a.Speak())'"
    exit 1
}
Write-Host "PASS"
exit 0
