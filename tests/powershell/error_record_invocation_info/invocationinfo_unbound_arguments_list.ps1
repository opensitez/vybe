# vybe-test: powershell/error_record_invocation_info/invocationinfo_unbound_arguments_list
function Throw-InfoCheck {
    param()
    throw "CheckInvocation"
}
$err = $null
try {
    Throw-InfoCheck
} catch {
    $err = $_
}
if ($err.InvocationInfo -eq $null) {
    Write-Host "FAIL: InvocationInfo missing"
    exit 1
}
Write-Host "PASS"
exit 0
