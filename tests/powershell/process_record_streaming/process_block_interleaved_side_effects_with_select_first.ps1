# vybe-test: powershell/process_record_streaming/process_block_interleaved_side_effects_with_select_first
# Downstream early termination with Select-Object -First halts upstream process block invocations immediately
$script:processedCounter = 0

function InfiniteStreamingSource {
    process {
        $script:processedCounter++
        $_
    }
}

$selected = 1..1000 | InfiniteStreamingSource | Select-Object -First 4

if ($selected.Count -ne 4) {
    Write-Host "FAIL: expected 4 selected items, got $($selected.Count)"
    exit 1
}

# The upstream function must have been invoked only 4 times, proving no unnecessary buffering
if ($script:processedCounter -ne 4) {
    Write-Host "FAIL: upstream process invoked $($script:processedCounter) times instead of 4"
    exit 1
}

Write-Host "PASS"
exit 0
