# vybe-test: powershell/join_string_cmdlet/join_string_empty_collection_yields_empty_string
$empty = @()

# Joining an empty collection produces an empty string, not $null
$res = $empty | Join-String -Separator ','

if ($null -eq $res) {
    Write-Host "FAIL: Join-String on empty collection returned `$null instead of empty string"
    exit 1
}

if ($res.Length -ne 0) {
    Write-Host "FAIL: expected length 0, got length $($res.Length): '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
