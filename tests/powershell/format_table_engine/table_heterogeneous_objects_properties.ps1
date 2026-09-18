# vybe-test: powershell/format_table_engine/table_heterogeneous_objects_properties
# Heterogeneous collection where objects have disjoint property sets
$hetero = @(
    [pscustomobject]@{ TypeA = "AlphaVal" },
    [pscustomobject]@{ TypeB = "BetaVal" }
)

$output = $hetero | Format-Table -Property TypeA, TypeB | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# Both properties must be formatted in their respective columns
if (-not ($output -match "AlphaVal" -and $output -match "BetaVal")) {
    Write-Host "FAIL: heterogeneous objects failed to format correctly: $output"
    exit 1
}

Write-Host "PASS"
exit 0
