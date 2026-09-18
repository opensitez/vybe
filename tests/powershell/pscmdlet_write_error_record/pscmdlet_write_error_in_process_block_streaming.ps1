# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_in_process_block_streaming
# Emitting non-terminating errors during streaming continues processing all subsequent pipeline items
function ProcessFilteringErrors {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        if ($InputObject % 2 -eq 1) {
            $ex = [System.ArgumentException]::new("Odd numbers not supported: $InputObject")
            $err = [System.Management.Automation.ErrorRecord]::new(
                $ex,
                "OddNumberId",
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $InputObject
            )
            $PSCmdlet.WriteError($err)
        } else {
            $InputObject
        }
    }
}

$results = @(1..4 | ProcessFilteringErrors 2>&1)

# Piped 1, 2, 3, 4:
# 1 -> ErrorRecord
# 2 -> 2
# 3 -> ErrorRecord
# 4 -> 4
if ($results.Count -ne 4) {
    Write-Host "FAIL: expected 4 items, got $($results.Count)"
    exit 1
}

$errors = @($results | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
$outputs = @($results | Where-Object { $_ -is [int] })

if ($errors.Count -ne 2 -or $outputs.Count -ne 2) {
    Write-Host "FAIL: error and output counts mismatch: errors=$($errors.Count), outputs=$($outputs.Count)"
    exit 1
}

if ($outputs[0] -ne 2 -or $outputs[1] -ne 4) {
    Write-Host "FAIL: outputs values mismatch: $($outputs[0]), $($outputs[1])"
    exit 1
}

Write-Host "PASS"
exit 0
