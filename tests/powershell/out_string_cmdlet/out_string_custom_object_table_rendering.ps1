# vybe-test: powershell/out_string_cmdlet/out_string_custom_object_table_rendering
# PSCustomObject converts into tabular string formatting with headers and dashed dividers
$obj = [pscustomobject]@{ MetricName = "Latency"; ValueMs = 12 }
$output = $obj | Out-String

if ($output -notmatch "MetricName" -or $output -notmatch "ValueMs") {
    Write-Host "FAIL: property column headers missing in Out-String table output"
    exit 1
}

if ($output -notmatch "Latency" -or $output -notmatch "12") {
    Write-Host "FAIL: property values missing in Out-String table output"
    exit 1
}

Write-Host "PASS"
exit 0
