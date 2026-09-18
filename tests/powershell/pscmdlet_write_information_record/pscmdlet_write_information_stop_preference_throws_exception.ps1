# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_stop_preference_throws_exception
# When $InformationPreference is Stop, calling $PSCmdlet.WriteInformation throws ActionPreferenceStopException
function TriggerStopInfo {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteInformation("terminating info message", @("StopTag"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Stop"

$threwStop = $false
try {
    TriggerStopInfo
} catch [System.Management.Automation.ActionPreferenceStopException] {
    $threwStop = $true
} finally {
    $InformationPreference = $oldPreference
}

if (-not $threwStop) {
    Write-Host "FAIL: WriteInformation under Stop preference did not throw ActionPreferenceStopException"
    exit 1
}

Write-Host "PASS"
exit 0
