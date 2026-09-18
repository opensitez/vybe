# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_in_end_block
# $PSCmdlet.WriteError emits an ErrorRecord from within the function end block
function EndErrorRoutine {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    begin { $count = 0 }
    process { $count++ }
    end {
        if ($count -lt 10) {
            $ex = [System.InvalidOperationException]::new("Insufficient samples processed: $count")
            $err = [System.Management.Automation.ErrorRecord]::new(
                $ex,
                "InsufficientSamplesId",
                [System.Management.Automation.ErrorCategory]::ResourceUnavailable,
                $count
            )
            $PSCmdlet.WriteError($err)
        }
    }
}

$captured = @(1..3 | EndErrorRoutine 2>&1)

if ($captured.Count -ne 1) {
    Write-Host "FAIL: expected 1 end block error, got $($captured.Count)"
    exit 1
}

if ($captured[0].Exception.Message -ne "Insufficient samples processed: 3") {
    Write-Host "FAIL: end error message mismatch: '$($captured[0].Exception.Message)'"
    exit 1
}

Write-Host "PASS"
exit 0
