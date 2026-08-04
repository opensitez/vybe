# vybe-test: powershell/regex/regex_match_groups
$str = "2024-07-15"
if ($str -match "(\d{4})-(\d{2})-(\d{2})") {
    if ($Matches[1] -ne "2024") { Write-Host "FAIL: year"; exit 1 }
    if ($Matches[2] -ne "07")   { Write-Host "FAIL: month"; exit 1 }
    if ($Matches[3] -ne "15")   { Write-Host "FAIL: day"; exit 1 }
} else {
    Write-Host "FAIL: no match"
    exit 1
}
Write-Host "PASS"
exit 0
