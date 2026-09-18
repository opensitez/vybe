# vybe-test: powershell/get_unique_cmdlet/get_unique_empty_collection_returns_null
# Piping an empty collection into Get-Unique evaluates to $null without throwing
$res = @() | Get-Unique

if ($null -ne $res) {
    Write-Host "FAIL: expected `$null from empty collection stream, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
