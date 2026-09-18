# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_record_message_property
# DebugRecord.Message contains the exact string payload supplied to $PSCmdlet.WriteDebug
function EmitMessageTest {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("Transaction #9842 verified successfully")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $record = (EmitMessageTest *>&1)[0]

    if ($record.Message -ne "Transaction #9842 verified successfully") {
        Write-Host "FAIL: message mismatch, got '$($record.Message)'"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
