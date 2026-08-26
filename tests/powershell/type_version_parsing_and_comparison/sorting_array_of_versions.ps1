# vybe-test: powershell/type_version_parsing_and_comparison/sorting_array_of_versions
$arr = @([version]"2.0", [version]"10.0", [version]"1.5", [version]"2.1")
$sorted = $arr | Sort-Object
if ($sorted[0] -ne [version]"1.5" -or $sorted[3] -ne [version]"10.0") {
    Write-Host "FAIL: Version sorting order incorrect"
    exit 1
}
Write-Host "PASS"
exit 0
