# vybe-test: powershell/intrinsic_collection_overloads/where_mode_last_no_count_returns_single
$source = 10..50
# .Where({ ... }, 'Last') without the count parameter returns a collection containing only the single last matching item
$res = $source.Where({ $_ % 7 -eq 0 }, 'Last')

if ($null -eq $res) {
    Write-Host "FAIL: .Where('Last') without count returned `$null"
    exit 1
}

if ($res.Count -ne 1) {
    Write-Host "FAIL: expected count 1, got $($res.Count)"
    exit 1
}

if ($res[0] -ne 49) {
    Write-Host "FAIL: expected last multiple of 7 up to 50 to be 49, got $($res[0])"
    exit 1
}

Write-Host "PASS"
exit 0
