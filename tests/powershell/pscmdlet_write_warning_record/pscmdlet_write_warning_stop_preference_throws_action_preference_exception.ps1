# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_stop_preference_throws_action_preference_exception
# When $WarningPreference is Stop, calling $PSCmdlet.WriteWarning throws ActionPreferenceStopException
function TriggerWarningStop {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("fatal warning threshold crossed")
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Stop"

$threwStop = $false
try {
    TriggerWarningStop
} catch [System.Management.Automation.ActionPreferenceStopException] {
    $threwStop = $true
} finally {
    $WarningPreference = $oldPreference
}

if (-not $threwStop) {
    Write-Host "FAIL: WriteWarning under Stop preference did not throw ActionPreferenceStopException"
    exit 1
}

Write-Host "PASS"
exit 0
