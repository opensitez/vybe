# vybe-test: powershell/type_version_parsing_and_comparison/comparison_greater_minor
$v1 = [version]"1.10.0"
$v2 = [version]"1.2.0"
if (-not ($v1 -gt $v2)) {
    Write-Host "FAIL: Minor version numeric comparison failed"
    exit 1
}
Write-Host "PASS"
exit 0
