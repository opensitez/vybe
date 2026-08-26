# vybe-test: powershell/parameters_validate_pattern/validatepattern_url_format
function Set-Url {
    param([ValidatePattern('^https?://')][string]$Url)
    return $Url
}
$res = Set-Url -Url "https://powershell.org"
if ($res -ne "https://powershell.org") {
    Write-Host "FAIL: ValidatePattern URL failed"
    exit 1
}
Write-Host "PASS"
exit 0
