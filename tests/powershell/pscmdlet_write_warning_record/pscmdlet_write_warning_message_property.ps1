# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_message_property
# WarningRecord.Message contains the exact message payload passed to $PSCmdlet.WriteWarning
function EmitWarningMessage {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("Database connection pool near capacity: 92% utilized")
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $record = (EmitWarningMessage 3>&1)[0]

    if ($record.Message -ne "Database connection pool near capacity: 92% utilized") {
        Write-Host "FAIL: Message mismatch, got '$($record.Message)'"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
