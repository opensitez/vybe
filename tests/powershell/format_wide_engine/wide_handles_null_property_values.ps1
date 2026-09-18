# vybe-test: powershell/format_wide_engine/wide_handles_null_property_values
$items = @(
    [pscustomobject]@{ Value = $null },
    [pscustomobject]@{ Value = "ValidWideText" }
)

# $null property values must render blank space without throwing exceptions
$output = $items | Format-Wide -Property Value -Column 2 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "ValidWideText")) {
    Write-Host "FAIL: non-null value failed to render: $output"
    exit 1
}

Write-Host "PASS"
exit 0
