# vybe-test: powershell/type_complex_numbers_arithmetic/complex_tostring_formatting
$c = [System.Numerics.Complex]::new(1.0, 2.0)
$str = $c.ToString()
if ($str -ne "<1; 2>" -and $str -ne "(1, 2)") {
    Write-Host "FAIL: Complex ToString failed, got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
