# vybe-test: powershell/type_char_classification_methods/char_range_generation
$range = [char]'a'..[char]'e'
if ($range.Length -ne 5 -or [char]$range[0] -ne [char]'a' -or [char]$range[4] -ne [char]'e') {
    Write-Host "FAIL: Char range generation failed"
    exit 1
}
Write-Host "PASS"
exit 0
