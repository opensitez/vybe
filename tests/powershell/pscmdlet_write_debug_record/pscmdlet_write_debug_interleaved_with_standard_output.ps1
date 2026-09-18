# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_interleaved_with_standard_output
# Emitting $PSCmdlet.WriteDebug alongside regular pipeline output preserves the relative interleaving of items
function InterleavedWorker {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        $PSCmdlet.WriteDebug("before:$InputObject")
        "out:$InputObject"
        $PSCmdlet.WriteDebug("after:$InputObject")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $stream = @(1 | InterleavedWorker *>&1)

    # Expected order: DebugRecord(before:1), String(out:1), DebugRecord(after:1)
    if ($stream.Count -ne 3) {
        Write-Host "FAIL: expected 3 stream items, got $($stream.Count)"
        exit 1
    }

    if ($stream[0].GetType().Name -ne "DebugRecord" -or $stream[0].Message -ne "before:1") {
        Write-Host "FAIL: first item was not before-debug"
        exit 1
    }

    if ($stream[1] -ne "out:1") {
        Write-Host "FAIL: second item was not standard output"
        exit 1
    }

    if ($stream[2].GetType().Name -ne "DebugRecord" -or $stream[2].Message -ne "after:1") {
        Write-Host "FAIL: third item was not after-debug"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
