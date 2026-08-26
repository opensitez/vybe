# vybe-test: powershell/regex_lookaround_assertions/password_validation_multiple_lookaheads
# Must have at least 1 digit, 1 lowercase, 1 uppercase, min 8 chars
$regex = "^(?=.*\d)(?=.*[a-z])(?=.*[A-Z]).{8,}$"
$good = "Pass1234" -match $regex
$bad = "password" -match $regex
if (-not $good -or $bad) {
    Write-Host "FAIL: Multiple lookahead password validation failed"
    exit 1
}
Write-Host "PASS"
exit 0
