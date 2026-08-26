# vybe-test: powershell/string_builder_operations/append_subarray_of_chars
[char[]]$chars = @([char]'a', [char]'b', [char]'c', [char]'d', [char]'e')
$sb = [System.Text.StringBuilder]::new()
$null = $sb.Append($chars, 1, 3) # b, c, d
if ($sb.ToString() -ne "bcd") {
    Write-Host "FAIL: Append char subarray failed"
    exit 1
}
Write-Host "PASS"
exit 0
