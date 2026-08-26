# vybe-test: powershell/string_builder_operations/append_strings
$sb = [System.Text.StringBuilder]::new()
$null = $sb.Append("Hello")
$null = $sb.Append(" ")
$null = $sb.Append("World")
if ($sb.ToString() -ne "Hello World" -or $sb.Length -ne 11) {
    Write-Host "FAIL: Append failed, got $($sb.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
