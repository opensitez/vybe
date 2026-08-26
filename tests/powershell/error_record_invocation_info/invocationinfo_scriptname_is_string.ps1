# vybe-test: powershell/error_record_invocation_info/invocationinfo_scriptname_is_string
$err = $null
try { throw "NameErr" } catch { $err = $_ }
if ($err.InvocationInfo.ScriptName -isnot [string]) {
    Write-Host "FAIL: ScriptName should be string"
    exit 1
}
Write-Host "PASS"
exit 0
