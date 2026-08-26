# vybe-test: powershell/error_record_invocation_info/invocationinfo_invoking_scriptblock
$sb = { throw "SbError" }
$err = $null
try {
    & $sb
} catch {
    $err = $_
}
if ($err.InvocationInfo -eq $null) {
    Write-Host "FAIL: ScriptBlock InvocationInfo missing"
    exit 1
}
Write-Host "PASS"
exit 0
