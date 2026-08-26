# vybe-test: powershell/type_version_parsing_and_comparison/undefined_component_returns_negative_one
$v = [version]"1.2"
if ($v.Build -ne -1 -or $v.Revision -ne -1) {
    Write-Host "FAIL: Undefined components should be -1"
    exit 1
}
Write-Host "PASS"
exit 0
