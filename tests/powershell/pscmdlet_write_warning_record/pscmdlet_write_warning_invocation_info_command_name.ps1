# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_invocation_info_command_name
# WarningRecord.InvocationInfo.MyCommand.Name matches the name of the function emitting the warning
function AssertWarningSourceCommand {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("command identity warning")
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $record = (AssertWarningSourceCommand 3>&1)[0]

    $cmdName = $record.InvocationInfo.MyCommand.Name
    if ($cmdName -ne "AssertWarningSourceCommand") {
        Write-Host "FAIL: expected MyCommand.Name 'AssertWarningSourceCommand', got '$cmdName'"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
