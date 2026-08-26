# vybe-test: powershell/parameters_validate_pattern/validatepattern_default_value_matches
function Get-Branch {
    param([ValidatePattern('^(main|dev|feature/.*)$')][string]$Branch = "main")
    return $Branch
}
$res = Get-Branch
if ($res -ne "main") {
    Write-Host "FAIL: ValidatePattern default value failed"
    exit 1
}
Write-Host "PASS"
exit 0
