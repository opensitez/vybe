# vybe-test: powershell/format_list_engine/list_group_by_property
$users = @(
    [pscustomobject]@{ Team = "Alpha"; User = "Alice" },
    [pscustomobject]@{ Team = "Beta";  User = "Bob" }
)

# -GroupBy segments vertical list items into distinct group blocks
$output = $users | Format-List -Property User -GroupBy Team | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# Group headers must appear
if (-not ($output -match "Team:\s*Alpha" -and $output -match "Team:\s*Beta")) {
    Write-Host "FAIL: group headers missing: $output"
    exit 1
}

# Group items must appear
if (-not ($output -match "User\s*:\s*Alice" -and $output -match "User\s*:\s*Bob")) {
    Write-Host "FAIL: grouped items missing: $output"
    exit 1
}

Write-Host "PASS"
exit 0
