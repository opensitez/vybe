# vybe-test: powershell/string_builder_operations/clear_buffer
$sb = [System.Text.StringBuilder]::new("data")
$null = $sb.Clear()
if ($sb.Length -ne 0 -or $sb.ToString() -ne "") {
    Write-Host "FAIL: Clear failed"
    exit 1
}
Write-Host "PASS"
exit 0
