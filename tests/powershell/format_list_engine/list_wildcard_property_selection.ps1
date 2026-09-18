# vybe-test: powershell/format_list_engine/list_wildcard_property_selection
$config = [pscustomobject]@{
    DbHost   = "localhost"
    DbPort   = 3306
    DbUser   = "root"
    HttpPort = 8080
}

# Wildcard Db* selects all database-related properties
$output = $config | Format-List -Property Db* | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# Matched properties must be present
if (-not ($output -match "DbHost\s*:\s*localhost" -and $output -match "DbPort\s*:\s*3306" -and $output -match "DbUser\s*:\s*root")) {
    Write-Host "FAIL: matched wildcard properties missing: $output"
    exit 1
}

# Non-matched property must NOT be present
if ($output -match "HttpPort") {
    Write-Host "FAIL: non-matched property 'HttpPort' unexpectedly present: $output"
    exit 1
}

Write-Host "PASS"
exit 0
