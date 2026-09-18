# vybe-test: powershell/format_list_engine/list_multiple_calculated_properties
$stats = [pscustomobject]@{ Hits = 1000; Errors = 50 }

$calcSuccess = @{ Label = "SuccessCount"; Expression = { $_.Hits - $_.Errors } }
$calcRate    = @{ Label = "ErrorRate";    Expression = { [math]::Round($_.Errors / $_.Hits, 3) } }

$output = $stats | Format-List $calcSuccess, $calcRate | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "SuccessCount\s*:\s*950" -and $output -match "ErrorRate\s*:\s*0\.05")) {
    Write-Host "FAIL: multiple calculated properties failed to render: $output"
    exit 1
}

Write-Host "PASS"
exit 0
