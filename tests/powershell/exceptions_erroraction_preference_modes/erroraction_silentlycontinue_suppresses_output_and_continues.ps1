# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_silentlycontinue_suppresses_output_and_continues
function Emit-Suppressed {
    [CmdletBinding()]
    param()
    Write-Error "SuppressedError"
    return "AfterSuppressed"
}
$res = Emit-Suppressed -ErrorAction SilentlyContinue
if ($res -ne "AfterSuppressed") {
    Write-Host "FAIL: -ErrorAction SilentlyContinue failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
