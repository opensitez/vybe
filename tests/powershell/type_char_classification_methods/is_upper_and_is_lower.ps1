# vybe-test: powershell/type_char_classification_methods/is_upper_and_is_lower
$u = [char]'H'
$l = [char]'h'
if (-not [char]::IsUpper($u) -or [char]::IsUpper($l) -or -not [char]::IsLower($l) -or [char]::IsLower($u)) {
    Write-Host "FAIL: IsUpper / IsLower check failed"
    exit 1
}
Write-Host "PASS"
exit 0
