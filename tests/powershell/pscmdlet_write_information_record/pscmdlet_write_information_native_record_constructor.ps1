# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_native_record_constructor
# Direct instantiation of InformationRecord via its constructor and passing to WriteInformation preserves fields
function EmitConstructedRecord {
    [CmdletBinding()]
    param()
    process {
        $rec = [System.Management.Automation.InformationRecord]::new(
            "ConstructedMessageData",
            "CustomSourceEngine"
        )
        $rec.Tags.Add("CustomTag1")
        $rec.Tags.Add("CustomTag2")
        $PSCmdlet.WriteInformation($rec)
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $record = (EmitConstructedRecord 6>&1)[0]

    if ($record.MessageData -ne "ConstructedMessageData") {
        Write-Host "FAIL: MessageData mismatch: '$($record.MessageData)'"
        exit 1
    }

    if ($record.Source -ne "CustomSourceEngine") {
        Write-Host "FAIL: Source mismatch: '$($record.Source)'"
        exit 1
    }

    if ($record.Tags.Count -ne 2 -or (-not $record.Tags.Contains("CustomTag2"))) {
        Write-Host "FAIL: Tags collection mismatch"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
