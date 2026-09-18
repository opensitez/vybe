# vybe-test: powershell/intrinsic_collection_overloads/where_mode_last_bounded_count
$source = 1..20
$res = $source.Where({ $_ % 2 -eq 0 }, 'Last', 3)

if ($null -eq $res) {
    Write-Host "FAIL: .Where('Last', 3) returned `$null"
    exit 1
}

if ($res.Count -ne 3) {
    Write-Host "FAIL: expected exactly 3 items, got $($res.Count)"
    exit 1
}

if ($res[0] -ne 16 -or $res[1] -ne 18 -or $res[2] -ne 20) {
    Write-Host "FAIL: expected values @(16, 18, 20), got @($($res -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
