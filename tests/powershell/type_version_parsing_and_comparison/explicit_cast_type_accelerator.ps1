# vybe-test: powershell/type_version_parsing_and_comparison/explicit_cast_type_accelerator
$v = [version]"1.0.0"
if ($v.GetType().Name -ne "Version") {
    Write-Host "FAIL: [version] cast failed"
    exit 1
}
Write-Host "PASS"
exit 0
