# vybe-test: powershell/readonly_variables/readonly_variable_hashtable_type
New-Variable -Name "RO_HASH" -Value @{ Mode = "Safe" } -Option ReadOnly
if ($RO_HASH.Mode -ne "Safe") {
    Write-Host "FAIL: ReadOnly hashtable expected Mode=Safe"
    exit 1
}
Write-Host "PASS"
exit 0
