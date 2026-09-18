# vybe-test: powershell/process_record_streaming/process_block_downstream_break_halts_upstream_process
# Executing a break statement in a downstream stage terminates the upstream pipeline stream
$script:upstreamIterations = 0

function InfiniteCounter {
    $n = 0
    while ($true) {
        $script:upstreamIterations++
        $n++
        $n
    }
}

$received = @()
InfiniteCounter | ForEach-Object {
    $received += $_
    if ($_ -ge 5) {
        break
    }
}

if ($received.Count -ne 5) {
    Write-Host "FAIL: expected 5 items collected before break, got $($received.Count)"
    exit 1
}

# Upstream should not continue producing infinitely after downstream break
if ($script:upstreamIterations -gt 6) {
    Write-Host "FAIL: upstream continued generating after downstream break: $($script:upstreamIterations) iterations"
    exit 1
}

Write-Host "PASS"
exit 0
