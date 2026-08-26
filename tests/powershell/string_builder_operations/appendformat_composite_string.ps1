# vybe-test: powershell/string_builder_operations/appendformat_composite_string
$sb = [System.Text.StringBuilder]::new()
$null = $sb.AppendFormat("User: {0}, ID: {1:D4}", "Alice", 42)
if ($sb.ToString() -ne "User: Alice, ID: 0042") {
    Write-Host "FAIL: AppendFormat failed, got $($sb.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
