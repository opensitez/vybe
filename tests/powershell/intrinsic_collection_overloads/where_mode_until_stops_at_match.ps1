# vybe-test: powershell/intrinsic_collection_overloads/where_mode_until_stops_at_match
$source = 1..10
# 'Until' collects elements until the predicate evaluates to $true, excluding the matching element itself
$res = $source.Where({ $_ -eq 5 }, 'Until')

if ($null -eq $res) {
    Write-Host "FAIL: .Where('Until') returned `$null"
    exit 1
}

if ($res.Count -ne 4) {
    Write-Host "FAIL: expected 4 items before 5, got $($res.Count)"
    exit 1
}

if ($res[0] -ne 1 -or $res[3] -ne 4) {
    Write-Host "FAIL: expected values @(1, 2, 3, 4), got @($($res -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
