# vybe-test: powershell/process_record_streaming/process_block_with_pipeline_parameter_by_value
# Declaring [Parameter(ValueFromPipeline=$true)] binds each pipeline object directly to the named parameter in process
function MultiplyByFactor {
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$Value,

        [int]$Factor = 3
    )
    process {
        $Value * $Factor
    }
}

$output = @(2, 4, 6) | MultiplyByFactor -Factor 5

if ($output.Count -ne 3) {
    Write-Host "FAIL: expected 3 output items, got $($output.Count)"
    exit 1
}

$expected = @(10, 20, 30)
for ($i = 0; $i -lt 3; $i++) {
    if ($output[$i] -ne $expected[$i]) {
        Write-Host "FAIL: at index ${i}, expected $($expected[$i]), got $($output[$i])"
        exit 1
    }
}

Write-Host "PASS"
exit 0
