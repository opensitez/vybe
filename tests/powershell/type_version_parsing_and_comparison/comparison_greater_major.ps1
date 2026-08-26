# vybe-test: powershell/type_version_parsing_and_comparison/comparison_greater_major
$v1 = [version]"3.0.0"
$v2 = [version]"2.9.9"
if (-not ($v1 -gt $v2)) {
    Write-Host "FAIL: Major version comparison failed"
    exit 1
}
Write-Host "PASS"
exit 0
