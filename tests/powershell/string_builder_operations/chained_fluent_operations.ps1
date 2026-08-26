# vybe-test: powershell/string_builder_operations/chained_fluent_operations
$sb = [System.Text.StringBuilder]::new()
$null = $sb.Append("start-").Append("mid-").Append("end")
if ($sb.ToString() -ne "start-mid-end") {
    Write-Host "FAIL: Chained append operations failed"
    exit 1
}
Write-Host "PASS"
exit 0
