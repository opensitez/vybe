# vybe-test: powershell/pipeline_sort_object_properties/sort_by_multiple_properties_mixed_order
$items = @(
    [pscustomobject]@{ Dept = "IT"; Salary = 5000 },
    [pscustomobject]@{ Dept = "HR"; Salary = 4000 },
    [pscustomobject]@{ Dept = "IT"; Salary = 7000 },
    [pscustomobject]@{ Dept = "HR"; Salary = 6000 }
)
$sorted = @($items | Sort-Object -Property @{ Expression = "Dept"; Descending = $false }, @{ Expression = "Salary"; Descending = $true })
if ($sorted[0].Dept -ne "HR" -or $sorted[0].Salary -ne 6000 -or $sorted[2].Dept -ne "IT" -or $sorted[2].Salary -ne 7000) {
    Write-Host "FAIL: Sort-Object multiple properties mixed order failed"
    exit 1
}
Write-Host "PASS"
exit 0
