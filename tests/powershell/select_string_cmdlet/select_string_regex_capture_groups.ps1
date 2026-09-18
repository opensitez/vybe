# vybe-test: powershell/select_string_cmdlet/select_string_regex_capture_groups
# MatchInfo.Matches exposes regex capture groups by 1-based index
$match = "Build status: SUCCEEDED in 45s" | Select-String -Pattern "status:\s*([A-Z]+)"

if ($null -eq $match) {
    Write-Host "FAIL: pattern match failed"
    exit 1
}

$capturedValue = $match.Matches[0].Groups[1].Value
if ($capturedValue -ne "SUCCEEDED") {
    Write-Host "FAIL: expected captured group value 'SUCCEEDED', got: '$capturedValue'"
    exit 1
}

Write-Host "PASS"
exit 0
