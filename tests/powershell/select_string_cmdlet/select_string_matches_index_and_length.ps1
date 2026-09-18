# vybe-test: powershell/select_string_cmdlet/select_string_matches_index_and_length
# MatchInfo.Matches exposes the start Index and Length of the matched substring
$line = "offset 12345 trailing"
$match = $line | Select-String -Pattern "[0-9]+"

if ($null -eq $match) {
    Write-Host "FAIL: pattern matching failed"
    exit 1
}

$m = $match.Matches[0]
if ($m.Index -ne 7) {
    Write-Host "FAIL: expected match start index 7, got: $($m.Index)"
    exit 1
}

if ($m.Length -ne 5) {
    Write-Host "FAIL: expected match length 5, got: $($m.Length)"
    exit 1
}

Write-Host "PASS"
exit 0
