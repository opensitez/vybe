# vybe-test: powershell/format_wide_engine/wide_property_selection
$items = @(
    [pscustomobject]@{ ItemName = "AlphaNode"; Secret = "TopSecret1" },
    [pscustomobject]@{ ItemName = "BetaNode";  Secret = "TopSecret2" }
)

# Format-Wide -Property selects the single property to display across columns
$output = $items | Format-Wide -Property ItemName | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# The selected property values must appear
if (-not ($output -match "AlphaNode" -and $output -match "BetaNode")) {
    Write-Host "FAIL: selected property 'ItemName' missing from output: $output"
    exit 1
}

# The unselected property values must NOT appear
if ($output -match "Secret" -or $output -match "TopSecret1" -or $output -match "TopSecret2") {
    Write-Host "FAIL: unselected property 'Secret' appeared in output: $output"
    exit 1
}

Write-Host "PASS"
exit 0
