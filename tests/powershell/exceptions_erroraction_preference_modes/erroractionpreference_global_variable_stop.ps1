# vybe-test: powershell/exceptions_erroraction_preference_modes/erroractionpreference_global_variable_stop
function Emit-GlobalCheck {
    [CmdletBinding()]
    param()
    Write-Error "GlobalStopCheck"
}
$oldEA = $ErrorActionPreference
$caught = $false
try {
    $ErrorActionPreference = "Stop"
    Emit-GlobalCheck
} catch {
    $caught = $true
} finally {
    $ErrorActionPreference = $oldEA
}
if (-not $caught) {
    Write-Host "FAIL: `$ErrorActionPreference='Stop' failed"
    exit 1
}
Write-Host "PASS"
exit 0
