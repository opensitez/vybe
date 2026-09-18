# vybe-test: powershell/process_record_streaming/process_block_pipeline_variable_common_parameter
# The -PipelineVariable common parameter stores each emitted object for access in subsequent pipeline stages
function EmitPrefixedValue {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$Number
    )
    process {
        $Number * 10
    }
}

$combined = @(1..3 | EmitPrefixedValue -PipelineVariable val | ForEach-Object { "$val-processed" })

if ($combined.Count -ne 3) {
    Write-Host "FAIL: expected 3 items, got $($combined.Count)"
    exit 1
}

$expected = "10-processed, 20-processed, 30-processed"
$actual = $combined -join ", "

if ($actual -ne $expected) {
    Write-Host "FAIL: expected '$expected', got '$actual'"
    exit 1
}

Write-Host "PASS"
exit 0
