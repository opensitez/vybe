# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_parameter_overrides_global_preference
function Emit-OverrideCheck {
    [CmdletBinding()]
    param()
    Write-Error "Overridden"
    return "Completed"
}
$oldEA = $ErrorActionPreference
try {
    $ErrorActionPreference = "Stop"
    # Local -ErrorAction SilentlyContinue overrides global Stop
    $res = Emit-OverrideCheck -ErrorAction SilentlyContinue
} finally {
    $ErrorActionPreference = $oldEA
}
if ($res -ne "Completed") {
    Write-Host "FAIL: Parameter -ErrorAction should override global `$ErrorActionPreference"
    exit 1
}
Write-Host "PASS"
exit 0
