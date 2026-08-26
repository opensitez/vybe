# vybe-test: powershell/type_version_parsing_and_comparison/tostring_truncated_field_count
$v = [version]"1.2.3.4"
$trunc = $v.ToString(2)
if ($trunc -ne "1.2") {
    Write-Host "FAIL: ToString(2) expected '1.2', got '$trunc'"
    exit 1
}
Write-Host "PASS"
exit 0
