# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_tags_collection
# Tags provided to $PSCmdlet.WriteInformation populate the InformationRecord.Tags collection
function EmitTaggedRecord {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteInformation("secure transaction", @("Security", "PCI_DSS", "Audit"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $record = (EmitTaggedRecord 6>&1)[0]

    if ($record.Tags.Count -ne 3) {
        Write-Host "FAIL: expected 3 tags, got $($record.Tags.Count)"
        exit 1
    }

    if (-not $record.Tags.Contains("PCI_DSS")) {
        Write-Host "FAIL: tag 'PCI_DSS' missing from Tags collection"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
