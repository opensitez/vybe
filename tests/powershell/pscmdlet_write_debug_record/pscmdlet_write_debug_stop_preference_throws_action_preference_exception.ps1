# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_stop_preference_throws_action_preference_exception
# When $DebugPreference is Stop, calling $PSCmdlet.WriteDebug throws ActionPreferenceStopException
function TriggerStopDebug {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("terminating debug signal")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Stop"

$threwStop = $false
try {
    TriggerStopDebug
} catch [System.Management.Automation.ActionPreferenceStopException] {
    $threwStop = $true
} finally {
    $DebugPreference = $oldPreference
}

if (-not $threwStop) {
    Write-Host "FAIL: $PSCmdlet.WriteDebug under Stop preference did not throw ActionPreferenceStopException"
    exit 1
}

Write-Host "PASS"
exit 0
