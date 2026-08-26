# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_with_null_value
$obj = [pscustomobject]@{ Unset = $null }
if ($obj.PSObject.Properties.Match("Unset").Count -ne 1 -or $obj.Unset -ne $null) {
    Write-Host "FAIL: PSNoteProperty with null value failed"
    exit 1
}
Write-Host "PASS"
exit 0
