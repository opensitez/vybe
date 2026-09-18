# vybe-test: powershell/format_wide_engine/wide_group_by_property
$animals = @(
    [pscustomobject]@{ Type = "Mammal"; Name = "Lion" },
    [pscustomobject]@{ Type = "Bird";   Name = "Eagle" }
)

# -GroupBy divides wide output into groups with distinct headers
$output = $animals | Format-Wide -Property Name -GroupBy Type | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# Group headers must appear
if (-not ($output -match "Type:\s*Mammal" -and $output -match "Type:\s*Bird")) {
    Write-Host "FAIL: GroupBy headers missing: $output"
    exit 1
}

# Items must appear
if (-not ($output -match "Lion" -and $output -match "Eagle")) {
    Write-Host "FAIL: items missing from grouped wide output: $output"
    exit 1
}

Write-Host "PASS"
exit 0
