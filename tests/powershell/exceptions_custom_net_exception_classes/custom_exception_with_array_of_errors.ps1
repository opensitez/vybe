# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_with_array_of_errors
class ValidationAggregateException : System.Exception {
    [string[]]$Errors
    ValidationAggregateException([string[]]$errs) : base("Multiple errors") {
        $this.Errors = $errs
    }
}
$vae = [ValidationAggregateException]::new(@("Error 1", "Error 2"))
if ($vae.Errors.Length -ne 2 -or $vae.Errors[0] -ne "Error 1") {
    Write-Host "FAIL: Custom exception with array of errors failed"
    exit 1
}
Write-Host "PASS"
exit 0
