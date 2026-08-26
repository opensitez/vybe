# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_inheriting_system_exception
class CustomAppException : System.Exception {
    CustomAppException() : base("Default app error") {}
    CustomAppException([string]$msg) : base($msg) {}
}
$ex = [CustomAppException]::new("Special error")
if ($ex.Message -ne "Special error" -or $ex -isnot [System.Exception]) {
    Write-Host "FAIL: Custom exception inheriting System.Exception failed"
    exit 1
}
Write-Host "PASS"
exit 0
