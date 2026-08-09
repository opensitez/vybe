# vybe-test: powershell/pscustomobject_literals/pscustomobject_case_insensitivity
$obj = [pscustomobject]@{ MixedCase = "Data" }
if ($obj.mixedcase -ne "Data") {
    Write-Host "FAIL: case-insensitive member access expected Data, got $($obj.mixedcase)"
    exit 1
}
Write-Host "PASS"
exit 0
