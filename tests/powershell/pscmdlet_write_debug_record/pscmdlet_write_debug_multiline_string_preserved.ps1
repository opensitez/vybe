# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_multiline_string_preserved
# Multi-line strings passed to $PSCmdlet.WriteDebug preserve embedded newlines in DebugRecord.Message
function EmitMultilineDebug {
    [CmdletBinding()]
    param()
    process {
        $msg = "Header: DebugTrace`nSubsystem: Network`nStatus: Connected"
        $PSCmdlet.WriteDebug($msg)
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $record = (EmitMultilineDebug *>&1)[0]

    $lines = $record.Message -split "`n"
    if ($lines.Count -ne 3) {
        Write-Host "FAIL: expected 3 lines in debug message, got $($lines.Count)"
        exit 1
    }

    if ($lines[0] -ne "Header: DebugTrace" -or $lines[2] -ne "Status: Connected") {
        Write-Host "FAIL: multiline content mismatch"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
