# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_ignore_preference_suppresses_record
# When $DebugPreference is set to Ignore, $PSCmdlet.WriteDebug output is completely bypassed
function CheckIgnorePreference {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("ignore stream entry")
        "retained output"
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Ignore"

try {
    $results = @(CheckIgnorePreference *>&1)

    if ($results.Count -ne 1) {
        Write-Host "FAIL: expected 1 item under Ignore, got $($results.Count)"
        exit 1
    }

    if ($results[0] -ne "retained output") {
        Write-Host "FAIL: output mismatch under Ignore preference"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
