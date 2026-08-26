# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_case_insensitivity
function Emit-CaseCheck {
    [CmdletBinding()]
    param()
    Write-Error "CaseErr"
}
$caught = $false
try {
    Emit-CaseCheck -ErrorAction "stop"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Case-insensitive -ErrorAction 'stop' failed"
    exit 1
}
Write-Host "PASS"
exit 0
