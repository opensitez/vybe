# vybe-test: powershell/format_wide_engine/wide_autosize_dynamic_width
$items = @(
    [pscustomobject]@{ Val = "Short" },
    [pscustomobject]@{ Val = "SomewhatLongerWideEntryValue" }
)

# -AutoSize calculates column spacing dynamically based on longest item
$output = $items | Format-Wide -Property Val -AutoSize | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "Short" -and $output -match "SomewhatLongerWideEntryValue")) {
    Write-Host "FAIL: AutoSize content missing: $output"
    exit 1
}

Write-Host "PASS"
exit 0
