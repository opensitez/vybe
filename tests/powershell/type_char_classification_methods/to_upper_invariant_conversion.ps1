# vybe-test: powershell/type_char_classification_methods/to_upper_invariant_conversion
$ch = [char]'k'
$up = [char]::ToUpperInvariant($ch)
if ($up -ne [char]'K') {
    Write-Host "FAIL: ToUpperInvariant expected K, got $up"
    exit 1
}
Write-Host "PASS"
exit 0
