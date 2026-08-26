# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_stop_turns_non_terminating_into_terminating
function Emit-WarningErr {
    [CmdletBinding()]
    param()
    Write-Error "NonTerminating"
}
$caught = $false
try {
    Emit-WarningErr -ErrorAction Stop
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: -ErrorAction Stop should convert non-terminating error to terminating exception"
    exit 1
}
Write-Host "PASS"
exit 0
