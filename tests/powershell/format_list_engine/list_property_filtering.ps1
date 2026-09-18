# vybe-test: powershell/format_list_engine/list_property_filtering
$person = [pscustomobject]@{
    Name   = "Alice"
    Age    = 30
    Role   = "Admin"
    Secret = "HiddenToken"
}

# Format-List -Property selects only specified properties into vertical list view
$output = $person | Format-List -Property Name, Role | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# Selected properties must be rendered
if (-not ($output -match "Name\s*:\s*Alice" -and $output -match "Role\s*:\s*Admin")) {
    Write-Host "FAIL: selected properties missing from list output: $output"
    exit 1
}

# Unselected properties must be omitted
if ($output -match "Secret" -or $output -match "HiddenToken" -or $output -match "Age") {
    Write-Host "FAIL: unselected properties appeared in list output: $output"
    exit 1
}

Write-Host "PASS"
exit 0
