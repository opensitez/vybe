# vybe-test: powershell/error_record_exception_chaining/targetinvocationexception_unwrapping
class TargetClass {
    static [void]Explode() { throw [System.InvalidOperationException]::new("TargetCrash") }
}
$err = $null
try {
    [TargetClass]::Explode()
} catch {
    $err = $_
}
if ($err.Exception.Message -ne "TargetCrash") {
    Write-Host "FAIL: Exception unwrapping failed, got '$($err.Exception.Message)'"
    exit 1
}
Write-Host "PASS"
exit 0
