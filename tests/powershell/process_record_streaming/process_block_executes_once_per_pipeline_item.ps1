# vybe-test: powershell/process_record_streaming/process_block_executes_once_per_pipeline_item
# A function with a process block executes the process block exactly once per piped pipeline object
$script:executionCount = 0

function CountProcessCalls {
    process {
        $script:executionCount++
    }
}

@(10, 20, 30, 40, 50) | CountProcessCalls

if ($script:executionCount -ne 5) {
    Write-Host "FAIL: expected 5 process executions, got $($script:executionCount)"
    exit 1
}

Write-Host "PASS"
exit 0
