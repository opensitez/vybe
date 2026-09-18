# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_variable_append_syntax
# Prefixing the variable name with '+' in -InformationVariable appends new records to existing collections
function StepEmitter {
    [CmdletBinding()]
    param([string]$StepName)
    process {
        $PSCmdlet.WriteInformation("Running $StepName", @("Workflow"))
    }
}

$history = @()
StepEmitter -StepName "StepA" -InformationVariable history -InformationAction SilentlyContinue
StepEmitter -StepName "StepB" -InformationVariable +history -InformationAction SilentlyContinue

if ($history.Count -ne 2) {
    Write-Host "FAIL: expected 2 appended records, got $($history.Count)"
    exit 1
}

if ($history[0].MessageData -ne "Running StepA" -or $history[1].MessageData -ne "Running StepB") {
    Write-Host "FAIL: appended record payload mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
