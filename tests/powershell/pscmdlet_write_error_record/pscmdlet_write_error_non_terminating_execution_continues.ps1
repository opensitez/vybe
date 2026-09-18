# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_non_terminating_execution_continues
# $PSCmdlet.WriteError is non-terminating by default; subsequent statements in the function execute
function NonTerminatingRoutine {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.InvalidOperationException]::new("Minor issue encountered")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "MinorIssueId",
            [System.Management.Automation.ErrorCategory]::InvalidOperation,
            $null
        )
        $PSCmdlet.WriteError($err)
        "execution survived after WriteError"
    }
}

$results = @(NonTerminatingRoutine 2>&1)

# Must contain ErrorRecord followed by String
if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 items, got $($results.Count)"
    exit 1
}

if ($results[1] -ne "execution survived after WriteError") {
    Write-Host "FAIL: statement after WriteError was not reached: '$($results[1])'"
    exit 1
}

Write-Host "PASS"
exit 0
