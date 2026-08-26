# vybe-test: powershell/type_version_parsing_and_comparison/constructor_four_components
$v = [version]::new(10, 0, 19041, 508)
if ($v.Major -ne 10 -or $v.Minor -ne 0 -or $v.Build -ne 19041 -or $v.Revision -ne 508) {
    Write-Host "FAIL: 4-component version constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
