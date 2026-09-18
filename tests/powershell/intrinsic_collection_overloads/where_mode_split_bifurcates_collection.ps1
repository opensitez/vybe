# vybe-test: powershell/intrinsic_collection_overloads/where_mode_split_bifurcates_collection
$source = 1..6
# 'Split' returns two collections: index 0 contains matching items, index 1 contains non-matching items
$split = $source.Where({ $_ % 2 -eq 0 }, 'Split')

if ($null -eq $split -or $split.Count -ne 2) {
    Write-Host "FAIL: .Where('Split') expected 2 partitions, got $($split.Count)"
    exit 1
}

$evens = $split[0]
$odds = $split[1]

if ($evens.Count -ne 3 -or $evens[0] -ne 2 -or $evens[2] -ne 6) {
    Write-Host "FAIL: matching partition expected @(2, 4, 6), got @($($evens -join ', '))"
    exit 1
}

if ($odds.Count -ne 3 -or $odds[0] -ne 1 -or $odds[2] -ne 5) {
    Write-Host "FAIL: non-matching partition expected @(1, 3, 5), got @($($odds -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
