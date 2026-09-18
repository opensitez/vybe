# vybe-test: powershell/intrinsic_collection_overloads/where_mode_skip_until_returns_remainder
$source = 1..10
# 'SkipUntil' discards elements until predicate is $true, then returns the match and all remaining items
$res = $source.Where({ $_ -eq 7 }, 'SkipUntil')

if ($null -eq $res) {
    Write-Host "FAIL: .Where('SkipUntil') returned `$null"
    exit 1
}

if ($res.Count -ne 4) {
    Write-Host "FAIL: expected 4 items starting from 7, got $($res.Count)"
    exit 1
}

if ($res[0] -ne 7 -or $res[3] -ne 10) {
    Write-Host "FAIL: expected values @(7, 8, 9, 10), got @($($res -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
