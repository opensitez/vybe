# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_with_array_value
$obj = [pscustomobject]@{ Items = @(1, 2, 3) }
$obj.Items += 4
if ($obj.Items.Length -ne 4 -or $obj.Items[3] -ne 4) {
    Write-Host "FAIL: PSNoteProperty with array value failed"
    exit 1
}
Write-Host "PASS"
exit 0
