# vybe-test: powershell/type_version_parsing_and_comparison/comparison_greater_build
$v1 = [version]"1.2.100"
$v2 = [version]"1.2.99"
if (-not ($v1 -gt $v2)) {
    Write-Host "FAIL: Build version comparison failed"
    exit 1
}
Write-Host "PASS"
exit 0
