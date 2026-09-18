# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_multiple_tags_matching
# Tag collections on InformationRecord support case-insensitive containment checks
function EmitTagsCheck {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteInformation("data payload", @("AlertLevelHigh", "RegionWest"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $record = (EmitTagsCheck 6>&1)[0]

    # Tags list in PowerShell provides case-insensitive matching
    $hasAlert = $record.Tags.Contains("AlertLevelHigh")
    $hasRegion = $record.Tags.Contains("RegionWest")
    $hasMissing = $record.Tags.Contains("NonExistentTag")

    if (-not $hasAlert -or (-not $hasRegion) -or $hasMissing) {
        Write-Host "FAIL: tag containment verification failed"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
