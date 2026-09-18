# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_variable_append_syntax
# Specifying +varName with -WarningVariable appends new WarningRecord instances across sequential executions
function StepWarningEmitter {
    [CmdletBinding()]
    param([string]$WarningText)
    process {
        $PSCmdlet.WriteWarning($WarningText)
    }
}

$warningList = @()
StepWarningEmitter -WarningText "Warning 1" -WarningVariable warningList -WarningAction SilentlyContinue
StepWarningEmitter -WarningText "Warning 2" -WarningVariable +warningList -WarningAction SilentlyContinue

if ($warningList.Count -ne 2) {
    Write-Host "FAIL: expected 2 accumulated warnings, got $($warningList.Count)"
    exit 1
}

if ($warningList[0].Message -ne "Warning 1" -or $warningList[1].Message -ne "Warning 2") {
    Write-Host "FAIL: accumulated warning messages mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
