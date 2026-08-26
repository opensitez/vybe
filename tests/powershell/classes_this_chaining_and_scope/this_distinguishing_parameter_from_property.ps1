# vybe-test: powershell/classes_this_chaining_and_scope/this_distinguishing_parameter_from_property
class UserProfile {
    [string]$Name
    [int]$Age
    [void]SetInfo([string]$Name, [int]$Age) {
        # Disambiguate local parameter from class field
        $this.Name = $Name
        $this.Age = $Age
    }
}
$u = [UserProfile]::new()
$u.SetInfo("Bob", 40)
if ($u.Name -ne "Bob" -or $u.Age -ne 40) {
    Write-Host "FAIL: Parameter shadow disambiguation with `$this failed"
    exit 1
}
Write-Host "PASS"
exit 0
