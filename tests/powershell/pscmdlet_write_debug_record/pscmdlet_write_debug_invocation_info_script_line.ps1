# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_invocation_info_script_line
# DebugRecord.InvocationInfo reflects the line information where the command emitting debug was invoked
function EmitWithInvocation {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("invocation test")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $record = (EmitWithInvocation *>&1)[0]

    if ($record.InvocationInfo -eq $null) {
        Write-Host "FAIL: InvocationInfo was null"
        exit 1
    }

    if ($record.InvocationInfo.ScriptLineNumber -le 0) {
        Write-Host "FAIL: ScriptLineNumber was not positive: $($record.InvocationInfo.ScriptLineNumber)"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
