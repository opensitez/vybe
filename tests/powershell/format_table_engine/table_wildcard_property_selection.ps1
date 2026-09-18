# vybe-test: powershell/format_table_engine/table_wildcard_property_selection
$device = [pscustomobject]@{
    NetIPAddress = "192.168.1.1"
    NetSubnet    = "255.255.255.0"
    DeviceName   = "Switch01"
}

# -Property with wildcards (Net*) selects only properties matching the pattern
$output = $device | Format-Table -Property Net* | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# The matched wildcard properties must appear
if (-not ($output -match "NetIPAddress" -and $output -match "NetSubnet")) {
    Write-Host "FAIL: matched wildcard properties missing from output: $output"
    exit 1
}

# The non-matching property must NOT appear
if ($output -match "DeviceName" -or $output -match "Switch01") {
    Write-Host "FAIL: non-matching property 'DeviceName' unexpectedly included: $output"
    exit 1
}

Write-Host "PASS"
exit 0
