# vybe-test: powershell/interpolation/subexpression_pipeline_interpolation
# A pipeline operation executed within $() expands its formatted result into the double-quoted string
$numbers = @(10, 25, 5, 40, 15)

$summary = "Filtered: $($numbers | Where-Object { $_ -ge 20 } | Sort-Object)"

# Filtered elements are 25 and 40
if ($summary -ne "Filtered: 25 40") {
    Write-Host "FAIL: expected 'Filtered: 25 40', got '$summary'"
    exit 1
}

Write-Host "PASS"
exit 0
