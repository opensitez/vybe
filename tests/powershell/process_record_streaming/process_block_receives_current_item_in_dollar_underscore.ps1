# vybe-test: powershell/process_record_streaming/process_block_receives_current_item_in_dollar_underscore
# Inside a process block, $_ holds the scalar value of the currently streaming pipeline object
$script:collected = @()

function RecordPipelineItems {
    process {
        $script:collected += $_
    }
}

@("alpha", "beta", "gamma") | RecordPipelineItems

if ($script:collected.Count -ne 3) {
    Write-Host "FAIL: expected 3 collected items, got $($script:collected.Count)"
    exit 1
}

$expectedJoined = "alpha, beta, gamma"
$actualJoined = $script:collected -join ", "
if ($actualJoined -ne $expectedJoined) {
    Write-Host "FAIL: expected '$expectedJoined', got '$actualJoined'"
    exit 1
}

Write-Host "PASS"
exit 0
