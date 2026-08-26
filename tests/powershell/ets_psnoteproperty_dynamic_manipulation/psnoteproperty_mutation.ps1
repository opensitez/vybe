# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_mutation
$obj = [pscustomobject]@{ Count = 1 }
$obj.Count = 100
if ($obj.Count -ne 100) {
    Write-Host "FAIL: PSNoteProperty mutation failed"
    exit 1
}
Write-Host "PASS"
exit 0
