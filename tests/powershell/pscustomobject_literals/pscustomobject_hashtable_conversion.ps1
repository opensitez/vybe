# vybe-test: powershell/pscustomobject_literals/pscustomobject_hashtable_conversion
$hash = @{ Key = "Val" }
$obj = [pscustomobject]$hash
if ($obj.Key -ne "Val") {
    Write-Host "FAIL: hashtable to PSCustomObject conversion expected Key=Val, got $($obj.Key)"
    exit 1
}
Write-Host "PASS"
exit 0
