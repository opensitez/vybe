# vybe-test: powershell/type_version_parsing_and_comparison/filtering_versions_greater_than
$arr = @([version]"1.0", [version]"2.0", [version]"3.0", [version]"4.0")
$filtered = $arr | Where-Object { $_ -ge [version]"2.5" }
if ($filtered.Count -ne 2 -or $filtered[0] -ne [version]"3.0") {
    Write-Host "FAIL: Filtered versions count or elements mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
