# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_scriptblock_execution_context
# An advanced function can invoke a scriptblock that calls $PSCmdlet.WriteDebug in the parent function context
function ExecuteWithCallerContext {
    [CmdletBinding()]
    param([scriptblock]$Action)
    process {
        & $Action
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $records = @(ExecuteWithCallerContext { $PSCmdlet.WriteDebug("scriptblock invoked debug") } *>&1)

    if ($records.Count -ne 1) {
        Write-Host "FAIL: expected 1 debug record from scriptblock invocation, got $($records.Count)"
        exit 1
    }

    if ($records[0].Message -ne "scriptblock invoked debug") {
        Write-Host "FAIL: message mismatch: '$($records[0].Message)'"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
