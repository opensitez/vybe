# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_ignore_preference_suppresses_record
# When $WarningPreference is set to Ignore, warning records are completely skipped
function EmitIgnoredWarning {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("ignored warning stream item")
        "remaining valid output"
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Ignore"

try {
    $results = @(EmitIgnoredWarning 3>&1)

    if ($results.Count -ne 1) {
        Write-Host "FAIL: expected 1 item under Ignore, got $($results.Count)"
        exit 1
    }

    if ($results[0] -ne "remaining valid output") {
        Write-Host "FAIL: output mismatch under Ignore preference"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
