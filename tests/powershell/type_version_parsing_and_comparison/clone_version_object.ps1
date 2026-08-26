# vybe-test: powershell/type_version_parsing_and_comparison/clone_version_object
$v1 = [version]"6.0.1"
$v2 = $v1.Clone()
if ($v1 -ne $v2 -or $v2.ToString() -ne "6.0.1") {
    Write-Host "FAIL: Version Clone failed"
    exit 1
}
Write-Host "PASS"
exit 0
