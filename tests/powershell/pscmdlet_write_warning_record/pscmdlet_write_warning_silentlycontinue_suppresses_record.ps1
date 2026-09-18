# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_silentlycontinue_suppresses_record
# When $WarningPreference is SilentlyContinue, $PSCmdlet.WriteWarning emits no records to stream 3
function EmitSilencedWarning {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("suppressed warning")
        "standard success payload"
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "SilentlyContinue"

try {
    $results = @(EmitSilencedWarning 3>&1)

    if ($results.Count -ne 1) {
        Write-Host "FAIL: expected 1 item, got $($results.Count)"
        exit 1
    }

    if ($results[0] -ne "standard success payload") {
        Write-Host "FAIL: unexpected stream output: $($results[0])"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
