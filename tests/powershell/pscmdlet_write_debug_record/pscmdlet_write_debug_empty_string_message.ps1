# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_empty_string_message
# $PSCmdlet.WriteDebug handles empty string inputs without error and produces a DebugRecord with empty message
function EmitEmptyDebug {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $records = @(EmitEmptyDebug *>&1)

    if ($records.Count -ne 1) {
        Write-Host "FAIL: expected 1 debug record, got $($records.Count)"
        exit 1
    }

    if ($records[0].Message -ne "") {
        Write-Host "FAIL: expected empty string message, got '$($records[0].Message)'"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
