# vybe-test: powershell/collections_keyvaluepair_struct/create_factory_method
$kvp = [System.Collections.Generic.KeyValuePair]::Create("name", "PowerShell")
if ($kvp.Key -ne "name" -or $kvp.Value -ne "PowerShell") {
    Write-Host "FAIL: KeyValuePair::Create factory method failed"
    exit 1
}
Write-Host "PASS"
exit 0
