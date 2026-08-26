# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_continue_default_writes_to_dollar_error
function Emit-Continued {
    [CmdletBinding()]
    param()
    Write-Error "ContinuedErrorRecord"
    return "ContinuedVal"
}
$res = Emit-Continued -ErrorAction Continue 2>$null
if ($res -ne "ContinuedVal" -or $Error[0].ToString() -notmatch "ContinuedErrorRecord") {
    Write-Host "FAIL: -ErrorAction Continue failed"
    exit 1
}
Write-Host "PASS"
exit 0
