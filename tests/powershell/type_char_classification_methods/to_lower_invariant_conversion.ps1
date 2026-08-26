# vybe-test: powershell/type_char_classification_methods/to_lower_invariant_conversion
$ch = [char]'M'
$low = [char]::ToLowerInvariant($ch)
if ($low -ne [char]'m') {
    Write-Host "FAIL: ToLowerInvariant expected m, got $low"
    exit 1
}
Write-Host "PASS"
exit 0
