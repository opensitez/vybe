# vybe-test: powershell/test_json_cmdlet/test_json_multiple_pipeline_items_streaming
$jsonStreams = @(
    '{"id": 1}',
    '{"id": 2}',
    '{"id": 3}'
)

# Piping multiple JSON strings yields a boolean result for each streamed item
$results = $jsonStreams | Test-Json

if ($results.Count -ne 3) {
    Write-Host "FAIL: expected 3 boolean results for 3 streamed items, got: $($results.Count)"
    exit 1
}

if ($results[0] -ne $true -or $results[1] -ne $true -or $results[2] -ne $true) {
    Write-Host "FAIL: not all streamed JSON documents returned `$true, got: @($($results -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
