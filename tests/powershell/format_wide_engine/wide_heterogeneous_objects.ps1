# vybe-test: powershell/format_wide_engine/wide_heterogeneous_objects
$items = @(
    [pscustomobject]@{ DisplayName = "DeviceOne" },
    [pscustomobject]@{ OtherProp   = "IgnoreMe" },
    [pscustomobject]@{ DisplayName = "DeviceTwo" }
)

# Format-Wide -Property extracts DisplayName where present, rendering blanks for missing
$output = $items | Format-Wide -Property DisplayName -Column 3 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "DeviceOne" -and $output -match "DeviceTwo")) {
    Write-Host "FAIL: heterogeneous objects failed in wide view: $output"
    exit 1
}

if ($output -match "IgnoreMe") {
    Write-Host "FAIL: unexpected property 'IgnoreMe' rendered in wide view: $output"
    exit 1
}

Write-Host "PASS"
exit 0
