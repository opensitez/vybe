# vybe-test: powershell/classes_hidden_members/hidden_property_type_enforcement
class TypedHidden {
    hidden [int]$Age
    SetAge([int]$a) { $this.Age = $a }
    [int]GetAge() { return $this.Age }
}
$t = [TypedHidden]::new()
$t.SetAge(25)
if ($t.GetAge() -ne 25) {
    Write-Host "FAIL: Hidden property type enforcement failed"
    exit 1
}
Write-Host "PASS"
exit 0
