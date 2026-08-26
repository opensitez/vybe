# vybe-test: powershell/string_builder_operations/tostring_substring_slice
$sb = [System.Text.StringBuilder]::new("0123456789")
$slice = $sb.ToString(3, 4)
if ($slice -ne "3456") {
    Write-Host "FAIL: ToString substring slice failed, got $slice"
    exit 1
}
Write-Host "PASS"
exit 0
