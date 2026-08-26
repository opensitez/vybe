# vybe-test: powershell/exceptions_custom_net_exception_classes/throwing_and_catching_custom_exception_by_type
class UserNotFoundException : System.Exception {
    [string]$Username
    UserNotFoundException([string]$user) : base("User not found: $user") {
        $this.Username = $user
    }
}
$caughtUser = ""
try {
    throw [UserNotFoundException]::new("alice")
} catch [UserNotFoundException] {
    $caughtUser = $_.Exception.Username
}
if ($caughtUser -ne "alice") {
    Write-Host "FAIL: Catching custom exception by type failed, got '$caughtUser'"
    exit 1
}
Write-Host "PASS"
exit 0
