# vybe-test: powershell/intrinsic_collection_overloads/where_mode_first_bounded_count
$source = 1..20
$res = $source.Where({ $_ % 2 -eq 0 }, 'First', 3)

if ($null -eq $res) {
    Write-Host "FAIL: .Where('First', 3) returned `$null"
    exit 1
}

if ($res.Count -ne 3) {
    Write-Host "FAIL: expected exactly 3 items, got $($res.Count)"
    exit 1
}

if ($res[0] -ne 2 -or $res[1] -ne 4 -or $res[2] -ne 6) {
    Write-Host "FAIL: expected values @(2, 4, 6), got @($($res -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
