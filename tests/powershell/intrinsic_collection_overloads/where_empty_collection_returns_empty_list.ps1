# vybe-test: powershell/intrinsic_collection_overloads/where_empty_collection_returns_empty_list
$emptyArray = @()

# .Where() on an empty collection must return a valid, non-null empty collection
$res = $emptyArray.Where({ $true })

if ($null -eq $res) {
    Write-Host "FAIL: .Where() on empty array returned `$null instead of empty collection"
    exit 1
}

if ($res.Count -ne 0) {
    Write-Host "FAIL: expected Count 0, got $($res.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
