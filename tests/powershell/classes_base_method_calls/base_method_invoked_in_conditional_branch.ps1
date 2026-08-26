# vybe-test: powershell/classes_base_method_calls/base_method_invoked_in_conditional_branch
class BaseCondition {
    [bool]IsAdmin() { return $true }
}
class SubSecurity : BaseCondition {
    [string]Check() {
        if (([BaseCondition]$this).IsAdmin()) { return "Allowed" }
        return "Denied"
    }
}
$sec = [SubSecurity]::new()
if ($sec.Check() -ne "Allowed") {
    Write-Host "FAIL: Base method in condition failed"
    exit 1
}
Write-Host "PASS"
exit 0
