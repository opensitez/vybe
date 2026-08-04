# vybe-test: powershell/classes/class_hidden_property
class Person {
    [string]$Name
    hidden [int]$Age
    Person([string]$n, [int]$a) { $this.Name = $n; $this.Age = $a }
    [string]Summary() { return "$($this.Name) is $($this.Age)" }
}
$p = [Person]::new("Alice", 30)
$s = $p.Summary()
if ($s -ne "Alice is 30") {
    Write-Host "FAIL: expected 'Alice is 30', got '$s'"
    exit 1
}
Write-Host "PASS"
exit 0
