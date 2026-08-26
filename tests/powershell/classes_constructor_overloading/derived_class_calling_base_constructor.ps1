# vybe-test: powershell/classes_constructor_overloading/derived_class_calling_base_constructor
class Animal {
    [string]$Species
    Animal([string]$s) { $this.Species = $s }
}
class Dog : Animal {
    [string]$Breed
    Dog([string]$b) : base("Canine") {
        $this.Breed = $b
    }
}
$d = [Dog]::new("Labrador")
if ($d.Species -ne "Canine" -or $d.Breed -ne "Labrador") {
    Write-Host "FAIL: Derived class calling base constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
