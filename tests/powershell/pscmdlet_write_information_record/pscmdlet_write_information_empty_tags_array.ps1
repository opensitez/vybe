# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_empty_tags_array
# Calling $PSCmdlet.WriteInformation with an empty tags collection produces an InformationRecord with 0 tags
function EmitEmptyTagsInfo {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteInformation("untagged payload", @())
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $record = (EmitEmptyTagsInfo 6>&1)[0]

    if ($record.Tags.Count -ne 0) {
        Write-Host "FAIL: expected 0 tags, got $($record.Tags.Count)"
        exit 1
    }

    if ($record.MessageData -ne "untagged payload") {
        Write-Host "FAIL: MessageData mismatch: '$($record.MessageData)'"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
