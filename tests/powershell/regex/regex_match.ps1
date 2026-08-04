# vybe-test: powershell/regex/regex_match
$str = "Hello World 123"
if ($str -match "\d+") {
    $num = $Matches[0]
    if ($num -ne "123") {
        Write-Host "FAIL: expected '123', got '$num'"
        exit 1
    }
} else {
    Write-Host "FAIL: match should have succeeded"
    exit 1
}
Write-Host "PASS"
exit 0
