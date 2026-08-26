# vybe-test: powershell/numeric_literal_forms/numeric_literal_forms_empty
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
