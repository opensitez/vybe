# vybe-test: powershell/process_record_streaming/process_block_yields_item_before_upstream_finishes
# Pipeline processing streams objects element-by-element rather than buffering the entire source array
$script:events = @()

function GenerateStream {
    for ($i = 1; $i -le 3; $i++) {
        $script:events += "generate:$i"
        $i
    }
}

function ConsumeStream {
    process {
        $script:events += "consume:$_"
        $_ * 10
    }
}

$results = @(GenerateStream | ConsumeStream)

# Interleaved event log demonstrates streaming execution order
$expectedLog = "generate:1, consume:1, generate:2, consume:2, generate:3, consume:3"
$actualLog = $script:events -join ", "

if ($actualLog -ne $expectedLog) {
    Write-Host "FAIL: events were not interleaved as streaming pipeline, got: '$actualLog'"
    exit 1
}

if ($results.Count -ne 3 -or $results[0] -ne 10 -or $results[2] -ne 30) {
    Write-Host "FAIL: unexpected stream result values"
    exit 1
}

Write-Host "PASS"
exit 0
