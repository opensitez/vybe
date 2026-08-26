# vybe-test: powershell/type_char_classification_methods/get_numeric_value_digits_and_fractions
$ch = [char]'8'
$val = [char]::GetNumericValue($ch)
if ($val -ne 8.0) {
    Write-Host "FAIL: GetNumericValue expected 8.0, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
