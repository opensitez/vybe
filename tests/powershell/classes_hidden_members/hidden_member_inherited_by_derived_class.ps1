# vybe-test: powershell/classes_hidden_members/hidden_member_inherited_by_derived_class
class BaseClass {
    hidden [string]$BaseSecret = "base_secret"
}
class SubClass : BaseClass {
    [string]Expose() { return $this.BaseSecret }
}
$s = [SubClass]::new()
if ($s.Expose() -ne "base_secret") {
    Write-Host "FAIL: Derived class accessing inherited hidden member failed"
    exit 1
}
Write-Host "PASS"
exit 0
