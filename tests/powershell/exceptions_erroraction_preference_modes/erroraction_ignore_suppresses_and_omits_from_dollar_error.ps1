# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_ignore_suppresses_and_omits_from_dollar_error
function Emit-Ignored {
    [CmdletBinding()]
    param()
    Write-Error "IgnoredErrorRecord"
}
$initialErrCount = $Error.Count
Emit-Ignored -ErrorAction Ignore
if ($Error.Count -ne $initialErrCount) {
    Write-Host "FAIL: -ErrorAction Ignore must not append to `$Error collection"
    exit 1
}
Write-Host "PASS"
exit 0
