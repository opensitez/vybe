# vybe-test: powershell/format_table_engine/table_property_column_filtering
$items = @(
    [pscustomobject]@{ Name = "Alice"; Age = 30; Secret = "HiddenToken" },
    [pscustomobject]@{ Name = "Bob";   Age = 25; Secret = "AnotherSecret" }
)

# Format-Table -Property filters only the selected columns into the output
$output = $items | Format-Table -Property Name, Age | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: Format-Table output was null"
    exit 1
}

# The selected properties must be present
if (-not ($output -match "Name" -and $output -match "Age" -and $output -match "Alice" -and $output -match "30")) {
    Write-Host "FAIL: selected properties 'Name' and 'Age' missing from output: $output"
    exit 1
}

# The unselected property must be omitted
if ($output -match "HiddenToken" -or $output -match "Secret") {
    Write-Host "FAIL: unselected property 'Secret' was unexpectedly included in output"
    exit 1
}

Write-Host "PASS"
exit 0
