# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_in_process_block_streams_per_item
# $PSCmdlet.WriteDebug inside process emits a debug record on each individual pipeline item
function ProcessStreamingDebug {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        $PSCmdlet.WriteDebug("processing item: $InputObject")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $records = @(10..12 | ProcessStreamingDebug *>&1)

    if ($records.Count -ne 3) {
        Write-Host "FAIL: expected 3 debug records, got $($records.Count)"
        exit 1
    }

    for ($i = 0; $i -lt 3; $i++) {
        $expectedMsg = "processing item: $(10 + $i)"
        if ($records[$i].Message -ne $expectedMsg) {
            Write-Host "FAIL: at index ${i}, expected '$expectedMsg', got '$($records[$i].Message)'"
            exit 1
        }
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
