# vybe-test: powershell/format_table_engine/table_handles_null_property_values
$data = @(
    [pscustomobject]@{ Key = "ValidKey"; Value = $null },
    [pscustomobject]@{ Key = "OtherKey"; Value = "NotNull" }
)

# Objects containing $null property values must render clean blank cells without throwing
$output = $data | Format-Table -Property Key, Value | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "ValidKey" -and $output -match "OtherKey" -and $output -match "NotNull")) {
    Write-Host "FAIL: table content with `$null values failed to render: $output"
    exit 1
}

Write-Host "PASS"
exit 0
