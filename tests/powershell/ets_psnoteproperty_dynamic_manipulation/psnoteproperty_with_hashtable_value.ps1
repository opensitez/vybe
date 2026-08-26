# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_with_hashtable_value
$obj = [pscustomobject]@{ Config = @{ env = "prod" } }
$obj.Config["port"] = 8080
if ($obj.Config["env"] -ne "prod" -or $obj.Config["port"] -ne 8080) {
    Write-Host "FAIL: PSNoteProperty with hashtable value failed"
    exit 1
}
Write-Host "PASS"
exit 0
