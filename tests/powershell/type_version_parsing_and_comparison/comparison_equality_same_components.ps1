# vybe-test: powershell/type_version_parsing_and_comparison/comparison_equality_same_components
$v1 = [version]"5.1.0"
$v2 = [version]"5.1.0"
if ($v1 -ne $v2) {
    Write-Host "FAIL: Identical versions must be equal"
    exit 1
}
Write-Host "PASS"
exit 0
