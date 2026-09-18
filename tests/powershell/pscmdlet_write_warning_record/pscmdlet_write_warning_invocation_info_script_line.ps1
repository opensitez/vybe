# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_invocation_info_script_line
# WarningRecord.InvocationInfo reflects the line information where the command was executed
function EmitWarningWithLine {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("line check warning")
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $record = (EmitWarningWithLine 3>&1)[0]

    if ($record.InvocationInfo -eq $null) {
        Write-Host "FAIL: InvocationInfo was null"
        exit 1
    }

    if ($record.InvocationInfo.ScriptLineNumber -le 0) {
        Write-Host "FAIL: ScriptLineNumber was not positive: $($record.InvocationInfo.ScriptLineNumber)"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
