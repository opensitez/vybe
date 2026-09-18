# vybe-test: powershell/format_list_engine/list_string_with_colons_in_value
$obj = [pscustomobject]@{
    Url = "https://example.com:8443/api/v1"
}

# Colons inside property values must not corrupt the Property : Value parsing or delimiter
$output = $obj | Format-List -Property Url | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "Url\s*:\s*https://example\.com:8443/api/v1")) {
    Write-Host "FAIL: URL with colons failed to format properly: $output"
    exit 1
}

Write-Host "PASS"
exit 0
