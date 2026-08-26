# vybe-test: powershell/type_version_parsing_and_comparison/tostring_full_components
$v = [version]"1.2.3.4"
if ($v.ToString() -ne "1.2.3.4") {
    Write-Host "FAIL: ToString() full components failed"
    exit 1
}
Write-Host "PASS"
exit 0
