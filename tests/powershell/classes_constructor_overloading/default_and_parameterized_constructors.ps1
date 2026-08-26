# vybe-test: powershell/classes_constructor_overloading/default_and_parameterized_constructors
class Person {
    [string]$Name
    [int]$Age
    Person() {
        $this.Name = "Unknown"
        $this.Age = 0
    }
    Person([string]$name, [int]$age) {
        $this.Name = $name
        $this.Age = $age
    }
}
$p1 = [Person]::new()
$p2 = [Person]::new("Alice", 30)
if ($p1.Name -ne "Unknown" -or $p1.Age -ne 0 -or $p2.Name -ne "Alice" -or $p2.Age -ne 30) {
    Write-Host "FAIL: Constructor overloading failed"
    exit 1
}
Write-Host "PASS"
exit 0
