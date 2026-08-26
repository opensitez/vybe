# vybe-test: powershell/string_culture_case_conversion/toupper_invariant
$str = "hello world"
$upper = $str.ToUpperInvariant()
if ($upper -ne "HELLO WORLD") {
    Write-Host "FAIL: ToUpperInvariant failed"
    exit 1
}
Write-Host "PASS"
exit 0
