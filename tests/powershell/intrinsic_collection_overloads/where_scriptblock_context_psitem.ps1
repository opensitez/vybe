# vybe-test: powershell/intrinsic_collection_overloads/where_scriptblock_context_psitem
$source = 1..10

# Within intrinsic .Where(), the current pipeline item is available as both $_ and $PSItem
$res = $source.Where({ $PSItem % 2 -ne 0 -and $PSItem -gt 5 })

if ($null -eq $res) {
    Write-Host "FAIL: .Where({ `$PSItem ... }) returned `$null"
    exit 1
}

if ($res.Count -ne 2) {
    Write-Host "FAIL: expected 2 odd items > 5 (7, 9), got $($res.Count)"
    exit 1
}

if ($res[0] -ne 7 -or $res[1] -ne 9) {
    Write-Host "FAIL: expected @(7, 9), got @($($res -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
