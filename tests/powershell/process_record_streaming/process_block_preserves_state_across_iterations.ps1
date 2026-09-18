# vybe-test: powershell/process_record_streaming/process_block_preserves_state_across_iterations
# Scope variables defined in a function persist across each consecutive process block iteration
function RunningSumAccumulator {
    begin {
        $runningSum = 0
    }
    process {
        $runningSum += $_
        $runningSum
    }
    end {
        # end block can verify total
    }
}

$runningSums = @(10, 20, 30, 40) | RunningSumAccumulator

if ($runningSums.Count -ne 4) {
    Write-Host "FAIL: expected 4 running sums, got $($runningSums.Count)"
    exit 1
}

$expected = @(10, 30, 60, 100)
for ($i = 0; $i -lt 4; $i++) {
    if ($runningSums[$i] -ne $expected[$i]) {
        Write-Host "FAIL: at index ${i}, expected $($expected[$i]), got $($runningSums[$i])"
        exit 1
    }
}

Write-Host "PASS"
exit 0
