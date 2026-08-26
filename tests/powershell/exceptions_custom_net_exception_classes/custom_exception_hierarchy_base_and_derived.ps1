# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_hierarchy_base_and_derived
class BaseDomainException : System.Exception {
    BaseDomainException([string]$m) : base($m) {}
}
class AccountLockedException : BaseDomainException {
    AccountLockedException([string]$m) : base($m) {}
}
$caught = $false
try {
    throw [AccountLockedException]::new("Locked out")
} catch [BaseDomainException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Catching derived custom exception via base type failed"
    exit 1
}
Write-Host "PASS"
exit 0
