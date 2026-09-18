# vybe-test: powershell/format_table_engine/table_format_enumeration_limit_ellipsis
$data = [pscustomobject]@{ Numbers = 1..10 }

# $FormatEnumerationLimit controls how many items of an inner collection are displayed before truncation
$prevLimit = $FormatEnumerationLimit
$FormatEnumerationLimit = 2

$output = $data | Format-Table -Property Numbers | Out-String

$FormatEnumerationLimit = $prevLimit

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# When limited to 2 items, the output displays {1, 2…} with ellipsis
$hasEllipsis = ($output -match "\u2026") -or ($output -match "\.\.\.")
if (-not $hasEllipsis) {
    Write-Host "FAIL: collection was not truncated with ellipsis under FormatEnumerationLimit=2, got: $output"
    exit 1
}

Write-Host "PASS"
exit 0
