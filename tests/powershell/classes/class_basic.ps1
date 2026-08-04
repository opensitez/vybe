# vybe-test: powershell/classes/class_basic
class Animal {
    [string]$Name
    Animal([string]$name) { $this.Name = $name }
    [string]Speak() { return "..." }
}
$a = [Animal]::new("Rex")
if ($a.Name -ne "Rex") {
    Write-Host "FAIL: expected 'Rex', got '$($a.Name)'"
    exit 1
}
Write-Host "PASS"
exit 0
