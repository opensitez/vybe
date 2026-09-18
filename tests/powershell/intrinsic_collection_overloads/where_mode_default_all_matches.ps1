# vybe-test: powershell/intrinsic_collection_overloads/where_mode_default_all_matches
$source = 1..10
# Explicit 'Default' mode collects all matching items across the entire collection
$res = $source.Where({ $_ % 3 -eq 0 }, 'Default')

if ($null -eq $res) {
    Write-Host "FAIL: .Where('Default') returned `$null"
    exit 1
}

if ($res.Count -ne 3) {
    Write-Host "FAIL: expected 3 matches (3, 6, 9), got $($res.Count)"
    exit 1
}

if ($res[0] -ne 3 -or $res[1] -ne 6 -or $res[2] -ne 9) {
    Write-Host "FAIL: expected values @(3, 6, 9), got @($($res -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
