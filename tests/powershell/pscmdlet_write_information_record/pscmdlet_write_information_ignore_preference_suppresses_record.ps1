# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_ignore_preference_suppresses_record
# When $InformationPreference is Ignore, InformationRecord output is completely bypassed even under 6>&1
function EmitIgnoredInfo {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteInformation("ignored info entry", @("IgnoreTag"))
        "persisted output"
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Ignore"

try {
    $results = @(EmitIgnoredInfo 6>&1)

    if ($results.Count -ne 1) {
        Write-Host "FAIL: expected 1 item under Ignore, got $($results.Count)"
        exit 1
    }

    if ($results[0] -ne "persisted output") {
        Write-Host "FAIL: unexpected result under Ignore preference"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
