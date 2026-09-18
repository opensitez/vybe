# vybe-test: powershell/intrinsic_collection_overloads/where_mode_first_no_count_returns_single
$source = 10..50
# .Where({ ... }, 'First') without the count parameter returns a collection containing only the single first matching item
$res = $source.Where({ $_ % 7 -eq 0 }, 'First')

if ($null -eq $res) {
    Write-Host "FAIL: .Where('First') without count returned `$null"
    exit 1
}

if ($res.Count -ne 1) {
    Write-Host "FAIL: expected count 1, got $($res.Count)"
    exit 1
}

if ($res[0] -ne 14) {
    Write-Host "FAIL: expected first multiple of 7 after 10 to be 14, got $($res[0])"
    exit 1
}

Write-Host "PASS"
exit 0
