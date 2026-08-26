# vybe-test: powershell/type_version_parsing_and_comparison/major_minor_build_revision_props
$v = [version]"4.5.6.7"
if ($v.Major -ne 4 -or $v.Minor -ne 5 -or $v.Build -ne 6 -or $v.Revision -ne 7) {
    Write-Host "FAIL: Property extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
