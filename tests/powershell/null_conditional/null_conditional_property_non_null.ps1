# vybe-test: powershell/null_conditional/null_conditional_property_non_null
$obj = [pscustomobject]@{ Name = "Vybe" }
$res = ${obj}?.Name
if ($res -ne "Vybe") {
    Write-Host "FAIL: non-null conditional property expected 'Vybe', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
