# vybe-test: powershell/parameters_validate_pattern/validatepattern_matching_email
function Set-UserEmail {
    param([ValidatePattern('^[\w\.-]+@[\w\.-]+\.\w+$')][string]$Email)
    return "Email:$Email"
}
$res = Set-UserEmail -Email "john.doe@company.org"
if ($res -ne "Email:john.doe@company.org") {
    Write-Host "FAIL: ValidatePattern email match failed"
    exit 1
}
Write-Host "PASS"
exit 0
