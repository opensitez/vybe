# vybe-test: powershell/type_version_parsing_and_comparison/comparison_greater_revision
$v1 = [version]"1.0.0.5"
$v2 = [version]"1.0.0.4"
if (-not ($v1 -gt $v2)) {
    Write-Host "FAIL: Revision comparison failed"
    exit 1
}
Write-Host "PASS"
exit 0
