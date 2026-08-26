# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_break_in_scriptblock_debugger
function Emit-ActionPrefCheck {
    [CmdletBinding()]
    param([System.Management.Automation.ActionPreference]$Mode)
    return $Mode
}
$res = Emit-ActionPrefCheck -Mode ([System.Management.Automation.ActionPreference]::Stop)
if ($res -ne [System.Management.Automation.ActionPreference]::Stop) {
    Write-Host "FAIL: ActionPreference parameter binding failed"
    exit 1
}
Write-Host "PASS"
exit 0
