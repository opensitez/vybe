# vybe-test: powershell/constant_variables/constant_variable_hashtable_type
New-Variable -Name "HASH_CONST" -Value @{ Active = $true } -Option Constant
if ($HASH_CONST.Active -ne $true) {
    Write-Host "FAIL: Constant hashtable expected Active=$true"
    exit 1
}
Write-Host "PASS"
exit 0
