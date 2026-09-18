# vybe-test: powershell/format_table_engine/table_direct_input_object_parameter
$singleObj = [pscustomobject]@{ MachineName = "AlphaNode"; CPU = 8 }

# Format-Table supports -InputObject explicitly outside of pipeline streaming
$output = Format-Table -InputObject $singleObj -Property MachineName, CPU | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "AlphaNode" -and $output -match "8" -and $output -match "MachineName")) {
    Write-Host "FAIL: direct -InputObject failed to format: $output"
    exit 1
}

Write-Host "PASS"
exit 0
