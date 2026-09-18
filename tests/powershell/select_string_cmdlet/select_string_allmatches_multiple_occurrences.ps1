# vybe-test: powershell/select_string_cmdlet/select_string_allmatches_multiple_occurrences
# -AllMatches captures every instance of a matching pattern on the line
$match = "code 101 status 200 error 404" | Select-String -Pattern "[0-9]{3}" -AllMatches

if ($null -eq $match) {
    Write-Host "FAIL: Select-String -AllMatches returned `$null"
    exit 1
}

if ($match.Matches.Count -ne 3) {
    Write-Host "FAIL: expected 3 regex matches, got: $($match.Matches.Count)"
    exit 1
}

$vals = @($match.Matches | ForEach-Object { $_.Value })
if ($vals[0] -ne "101" -or $vals[1] -ne "200" -or $vals[2] -ne "404") {
    Write-Host "FAIL: captured values mismatch: @($($vals -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
