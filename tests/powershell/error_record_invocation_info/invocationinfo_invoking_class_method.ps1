# vybe-test: powershell/error_record_invocation_info/invocationinfo_invoking_class_method
class MethodThrower {
    [void]Crash() { throw "ClassCrash" }
}
$mt = [MethodThrower]::new()
$err = $null
try {
    $mt.Crash()
} catch {
    $err = $_
}
if ($err.InvocationInfo -eq $null) {
    Write-Host "FAIL: Class method InvocationInfo missing"
    exit 1
}
Write-Host "PASS"
exit 0
