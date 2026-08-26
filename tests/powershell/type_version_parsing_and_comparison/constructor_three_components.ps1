# vybe-test: powershell/type_version_parsing_and_comparison/constructor_three_components
$v = [version]::new(1, 2, 3)
if ($v.Major -ne 1 -or $v.Minor -ne 2 -or $v.Build -ne 3 -or $v.Revision -ne -1) {
    Write-Host "FAIL: 3-component version constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
