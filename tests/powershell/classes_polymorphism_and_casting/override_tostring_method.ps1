# vybe-test: powershell/classes_polymorphism_and_casting/override_tostring_method
class CustomObject {
    [string]$Name
    CustomObject([string]$n) { $this.Name = $n }
    [string]ToString() { return "CustomObject($($this.Name))" }
}
$co = [CustomObject]::new("Alpha")
if ($co.ToString() -ne "CustomObject(Alpha)" -or "$co" -ne "CustomObject(Alpha)") {
    Write-Host "FAIL: ToString override failed, got '$co'"
    exit 1
}
Write-Host "PASS"
exit 0
