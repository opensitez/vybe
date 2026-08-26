# vybe-test: powershell/string_builder_operations/insert_substring_at_index
$sb = [System.Text.StringBuilder]::new("ac")
$null = $sb.Insert(1, "b")
if ($sb.ToString() -ne "abc") {
    Write-Host "FAIL: Insert failed, got $($sb.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
