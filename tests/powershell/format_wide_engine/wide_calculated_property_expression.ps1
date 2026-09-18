# vybe-test: powershell/format_wide_engine/wide_calculated_property_expression
$users = @(
    [pscustomobject]@{ Username = "alice" },
    [pscustomobject]@{ Username = "bob" }
)

# Format-Wide accepts a calculated property hashtable with Expression
$calc = @{ Expression = { $_.Username.ToUpper() } }
$output = $users | Format-Wide $calc -Column 2 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "ALICE" -and $output -match "BOB")) {
    Write-Host "FAIL: calculated property failed to render in wide view: $output"
    exit 1
}

Write-Host "PASS"
exit 0
