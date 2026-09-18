# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_in_end_block
# $PSCmdlet.WriteDebug emits debug records cleanly from within the end block
function EndDebugFunc {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    begin { $sum = 0 }
    process { $sum += $InputObject }
    end {
        $PSCmdlet.WriteDebug("final accumulated sum: $sum")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $records = @(1..4 | EndDebugFunc *>&1)

    if ($records.Count -ne 1) {
        Write-Host "FAIL: expected 1 end block debug record, got $($records.Count)"
        exit 1
    }

    if ($records[0].Message -ne "final accumulated sum: 10") {
        Write-Host "FAIL: message mismatch: '$($records[0].Message)'"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
