# vybe-test: powershell/classes_base_method_calls/base_method_returning_boolean
class BaseValidator {
    [bool]IsValid([string]$str) { return ($str.Length -gt 3) }
}
class SubValidator : BaseValidator {
    [bool]CheckString([string]$str) {
        return ([BaseValidator]$this).IsValid($str)
    }
}
$sv = [SubValidator]::new()
if (-not $sv.CheckString("valid") -or $sv.CheckString("no")) {
    Write-Host "FAIL: Base method boolean check failed"
    exit 1
}
Write-Host "PASS"
exit 0
