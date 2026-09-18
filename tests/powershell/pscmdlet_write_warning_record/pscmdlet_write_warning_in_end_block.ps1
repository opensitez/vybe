# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_in_end_block
# $PSCmdlet.WriteWarning emits WarningRecord entries from within the function end block
function EndWarningEmitter {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    begin { $count = 0 }
    process { $count++ }
    end {
        if ($count -gt 2) {
            $PSCmdlet.WriteWarning("Processed item count exceeds recommended threshold: $count")
        }
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $warnings = @(1..5 | EndWarningEmitter 3>&1)

    if ($warnings.Count -ne 1) {
        Write-Host "FAIL: expected 1 end block warning record, got $($warnings.Count)"
        exit 1
    }

    if ($warnings[0].Message -ne "Processed item count exceeds recommended threshold: 5") {
        Write-Host "FAIL: end warning message mismatch: '$($warnings[0].Message)'"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
