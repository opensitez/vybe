# vybe-test: powershell/string_builder_operations/length_expansion_with_null_chars
$sb = [System.Text.StringBuilder]::new("A")
$sb.Length = 3
if ($sb.Length -ne 3 -or $sb[0] -ne [char]'A' -or $sb[1] -ne [char]0) {
    Write-Host "FAIL: Length expansion with null chars failed"
    exit 1
}
Write-Host "PASS"
exit 0
