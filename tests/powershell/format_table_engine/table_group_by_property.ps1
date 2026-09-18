# vybe-test: powershell/format_table_engine/table_group_by_property
$staff = @(
    [pscustomobject]@{ Dept = "Engineering"; Name = "Alex" },
    [pscustomobject]@{ Dept = "Marketing";   Name = "Blake" }
)

# -GroupBy segments table output into groups with distinct headers
$output = $staff | Format-Table -Property Name -GroupBy Dept | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# The group headers must appear
if (-not ($output -match "Dept:\s*Engineering" -and $output -match "Dept:\s*Marketing")) {
    Write-Host "FAIL: GroupBy headers missing from output: $output"
    exit 1
}

# The row items must appear
if (-not ($output -match "Alex" -and $output -match "Blake")) {
    Write-Host "FAIL: grouped items missing from output: $output"
    exit 1
}

Write-Host "PASS"
exit 0
