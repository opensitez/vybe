# vybe-test: powershell/out_string_cmdlet/out_string_guid_table_formatting
# Guid instances format through Out-String with standard Guid column header
$guid = [Guid]::Parse("12345678-1234-1234-1234-123456789012")
$output = $guid | Out-String

if ($output -notmatch "Guid") {
    Write-Host "FAIL: 'Guid' column header missing from Guid Out-String formatting"
    exit 1
}

if ($output -notmatch "12345678-1234-1234-1234-123456789012") {
    Write-Host "FAIL: Guid value missing from Out-String formatting"
    exit 1
}

Write-Host "PASS"
exit 0
