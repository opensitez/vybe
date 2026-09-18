# vybe-test: powershell/intrinsic_collection_overloads/where_no_matches_returns_empty_collection
$source = 1..10

# When no elements satisfy the predicate, .Where() returns an empty collection, not $null
$res = $source.Where({ $_ -gt 500 })

if ($null -eq $res) {
    Write-Host "FAIL: .Where() with no matches returned `$null instead of empty collection"
    exit 1
}

if ($res.Count -ne 0) {
    Write-Host "FAIL: expected Count 0 on no matches, got $($res.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
