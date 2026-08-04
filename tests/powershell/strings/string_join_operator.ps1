# vybe-test: powershell/strings/string_join_array
$words = @("the", "quick", "brown", "fox")
$joined = $words -join " "
if ($joined -ne "the quick brown fox") {
    Write-Host "FAIL: '$joined'"
    exit 1
}
$csv = $words -join ","
if ($csv -ne "the,quick,brown,fox") {
    Write-Host "FAIL: csv '$csv'"
    exit 1
}
Write-Host "PASS"
exit 0
