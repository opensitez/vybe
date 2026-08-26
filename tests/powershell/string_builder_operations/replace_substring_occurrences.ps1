# vybe-test: powershell/string_builder_operations/replace_substring_occurrences
$sb = [System.Text.StringBuilder]::new("foo bar foo")
$null = $sb.Replace("foo", "baz")
if ($sb.ToString() -ne "baz bar baz") {
    Write-Host "FAIL: Replace failed, got $($sb.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
