# vybe-test: powershell/string_builder_operations/append_char_repeated_count
$sb = [System.Text.StringBuilder]::new()
$null = $sb.Append([char]'*', 5)
if ($sb.ToString() -ne "*****") {
    Write-Host "FAIL: Append repeated char failed"
    exit 1
}
Write-Host "PASS"
exit 0
