# vybe-test: powershell/string_builder_operations/remove_range_of_characters
$sb = [System.Text.StringBuilder]::new("abcdef")
$null = $sb.Remove(2, 3) # remove c, d, e
if ($sb.ToString() -ne "abf") {
    Write-Host "FAIL: Remove range failed, got $($sb.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
