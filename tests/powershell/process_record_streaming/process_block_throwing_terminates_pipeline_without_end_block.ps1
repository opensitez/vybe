# vybe-test: powershell/process_record_streaming/process_block_throwing_terminates_pipeline_without_end_block
# A terminating exception thrown during a process block halts further processing and skips the end block
$script:endBlockExecuted = $false
$script:itemsSeen = 0

function FaultyProcessor {
    process {
        $script:itemsSeen++
        if ($_ -eq 3) {
            throw "Intentional process failure at item 3"
        }
    }
    end {
        $script:endBlockExecuted = $true
    }
}

$threw = $false
try {
    1..5 | FaultyProcessor
} catch {
    $threw = $true
}

if (-not $threw) {
    Write-Host "FAIL: pipeline did not throw exception"
    exit 1
}

if ($script:itemsSeen -ne 3) {
    Write-Host "FAIL: expected 3 items seen before throw, got $($script:itemsSeen)"
    exit 1
}

if ($script:endBlockExecuted) {
    Write-Host "FAIL: end block ran despite uncaught exception in process block"
    exit 1
}

Write-Host "PASS"
exit 0
