# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_invocation_info_my_command_name
# DebugRecord.InvocationInfo.MyCommand.Name accurately reflects the advanced function name
function InvokeSpecificDebugCommand {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("ident check")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $record = (InvokeSpecificDebugCommand *>&1)[0]

    $cmdName = $record.InvocationInfo.MyCommand.Name
    if ($cmdName -ne "InvokeSpecificDebugCommand") {
        Write-Host "FAIL: expected MyCommand.Name 'InvokeSpecificDebugCommand', got '$cmdName'"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
