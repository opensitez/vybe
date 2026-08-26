# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_short_alias_ea
function Emit-AliasCheck {
    [CmdletBinding()]
    param()
    Write-Error "AliasError"
}
$caught = $false
try {
    Emit-AliasCheck -EA Stop
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: -EA Stop parameter alias failed"
    exit 1
}
Write-Host "PASS"
exit 0
