# vybe-test: powershell/string_builder_operations/replace_range_with_char
$sb = [System.Text.StringBuilder]::new("banana")
$null = $sb.Replace([char]'a', [char]'o', 1, 3) # replaces 'a' in "ana"
if ($sb.ToString() -ne "bonona") {
    Write-Host "FAIL: Replace range failed, got $($sb.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
