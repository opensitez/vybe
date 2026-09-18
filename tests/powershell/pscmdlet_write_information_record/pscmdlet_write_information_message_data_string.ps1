# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_message_data_string
# InformationRecord.MessageData preserves string messages emitted via $PSCmdlet.WriteInformation
function EmitStringInfo {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteInformation("Deployment batch #42 completed successfully", @("Deployment"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $record = (EmitStringInfo 6>&1)[0]

    if ($record.MessageData -ne "Deployment batch #42 completed successfully") {
        Write-Host "FAIL: MessageData mismatch, got '$($record.MessageData)'"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
