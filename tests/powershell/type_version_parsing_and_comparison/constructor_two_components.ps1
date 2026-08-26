# vybe-test: powershell/type_version_parsing_and_comparison/constructor_two_components
$v = [version]::new(2, 5)
if ($v.Major -ne 2 -or $v.Minor -ne 5 -or $v.Build -ne -1 -or $v.Revision -ne -1) {
    Write-Host "FAIL: 2-component version constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
