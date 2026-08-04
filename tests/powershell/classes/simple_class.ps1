# vybe-test: powershell/classes/simple_class
class Person {
    [string]$Name
    [int]$Age
    
    Person([string]$name, [int]$age) {
        $this.Name = $name
        $this.Age = $age
    }
    
    [string] Greet() {
        return "Hello, my name is $($this.Name)"
    }
}

$person = [Person]::new("Alice", 30)
if ($person.Name -ne "Alice") {
    Write-Host "FAIL: expected Name 'Alice', got '$($person.Name)'"
    exit 1
}
$greeting = $person.Greet()
if ($greeting -ne "Hello, my name is Alice") {
    Write-Host "FAIL: expected 'Hello, my name is Alice', got '$greeting'"
    exit 1
}
Write-Host "PASS"
exit 0
