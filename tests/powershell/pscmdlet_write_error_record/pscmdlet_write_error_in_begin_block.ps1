# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_in_begin_block
# Emitting a non-terminating error in the begin block allows execution to proceed into the process block
function BeginErrorRoutine {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    begin {
        $ex = [System.Exception]::new("Warning in begin initialization")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "BeginWarningId",
            [System.Management.Automation.ErrorCategory]::InvalidOperation,
            $null
        )
        $PSCmdlet.WriteError($err)
    }
    process {
        $InputObject * 10
    }
}

$results = @(1..2 | BeginErrorRoutine 2>&1)

# Expected: 1 ErrorRecord + 2 output numbers
if ($results.Count -ne 3) {
    Write-Host "FAIL: expected 3 items, got $($results.Count)"
    exit 1
}

if ($results[0].GetType().Name -ne "ErrorRecord") {
    Write-Host "FAIL: first item was not ErrorRecord"
    exit 1
}

if ($results[1] -ne 10 -or $results[2] -ne 20) {
    Write-Host "FAIL: process block outputs mismatch: $($results[1]), $($results[2])"
    exit 1
}

Write-Host "PASS"
exit 0
