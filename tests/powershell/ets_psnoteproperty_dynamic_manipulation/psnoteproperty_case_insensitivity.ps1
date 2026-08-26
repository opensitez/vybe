# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_case_insensitivity
$obj = [pscustomobject]@{ Target = "value" }
if ($obj.target -ne "value" -or $obj.TARGET -ne "value") {
    Write-Host "FAIL: PSNoteProperty case-insensitivity failed"
    exit 1
}
Write-Host "PASS"
exit 0
