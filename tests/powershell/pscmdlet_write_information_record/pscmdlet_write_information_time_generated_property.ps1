# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_time_generated_property
# InformationRecord.TimeGenerated contains a valid System.DateTime timestamp near the current system time
function EmitTimestampRecord {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteInformation("timestamp check", @("Test"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $before = [System.DateTime]::UtcNow.AddSeconds(-5)
    $record = (EmitTimestampRecord 6>&1)[0]
    $after = [System.DateTime]::UtcNow.AddSeconds(5)

    if ($record.TimeGenerated.GetType().Name -ne "DateTime") {
        Write-Host "FAIL: TimeGenerated is not DateTime"
        exit 1
    }

    $timeUtc = $record.TimeGenerated.ToUniversalTime()
    if ($timeUtc -lt $before -or $timeUtc -gt $after) {
        Write-Host "FAIL: TimeGenerated $timeUtc is outside expected window [$before, $after]"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
