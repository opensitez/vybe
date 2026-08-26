# vybe-test: powershell/string_builder_operations/appendjoin_collection
$sb = [System.Text.StringBuilder]::new()
$null = $sb.AppendJoin(", ", @("one", "two", "three"))
if ($sb.ToString() -ne "one, two, three") {
    Write-Host "FAIL: AppendJoin failed, got $($sb.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
