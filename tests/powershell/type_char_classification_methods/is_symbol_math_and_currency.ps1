# vybe-test: powershell/type_char_classification_methods/is_symbol_math_and_currency
$plus = [char]'+'
$dollar = [char]'$'
$num = [char]'1'
if (-not [char]::IsSymbol($plus) -or -not [char]::IsSymbol($dollar) -or [char]::IsSymbol($num)) {
    Write-Host "FAIL: IsSymbol check failed"
    exit 1
}
Write-Host "PASS"
exit 0
