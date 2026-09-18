# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_preserves_error_record_in_exception_on_stop
# When Stop preference triggers an exception, the exception's ErrorRecord retains the debug message text
function TriggerHalt {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("Critical debug checkpoint failure")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Stop"

$caughtMessage = ""
try {
    TriggerHalt
} catch [System.Management.Automation.ActionPreferenceStopException] {
    $caughtMessage = $_.Exception.Message
} finally {
    $DebugPreference = $oldPreference
}

if (-not $caughtMessage.Contains("Critical debug checkpoint failure")) {
    Write-Host "FAIL: exception message did not contain debug payload: '$caughtMessage'"
    exit 1
}

Write-Host "PASS"
exit 0
