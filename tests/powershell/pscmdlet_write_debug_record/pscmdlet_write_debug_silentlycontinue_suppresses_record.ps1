# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_silentlycontinue_suppresses_record
# When $DebugPreference is SilentlyContinue, $PSCmdlet.WriteDebug emits no records to the pipeline
function CheckSilentlyContinue {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("this should be silenced")
        "regular data"
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "SilentlyContinue"

try {
    $results = @(CheckSilentlyContinue *>&1)

    if ($results.Count -ne 1) {
        Write-Host "FAIL: expected 1 result, got $($results.Count)"
        exit 1
    }

    if ($results[0] -ne "regular data") {
        Write-Host "FAIL: unexpected stream output: $($results[0])"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
