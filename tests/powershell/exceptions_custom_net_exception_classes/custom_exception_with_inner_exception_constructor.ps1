# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_with_inner_exception_constructor
class WrappedException : System.Exception {
    WrappedException([string]$msg, [System.Exception]$inner) : base($msg, $inner) {}
}
$inner = [System.FormatException]::new("Invalid format")
$wrapped = [WrappedException]::new("Wrapper error", $inner)
if ($wrapped.InnerException -ne $inner -or $wrapped.Message -ne "Wrapper error") {
    Write-Host "FAIL: Custom exception with inner exception constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
