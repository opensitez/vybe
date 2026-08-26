# vybe-test: powershell/string_builder_operations/length_truncate
$sb = [System.Text.StringBuilder]::new("Hello World")
$sb.Length = 5
if ($sb.ToString() -ne "Hello") {
    Write-Host "FAIL: Length truncation failed, got $($sb.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
