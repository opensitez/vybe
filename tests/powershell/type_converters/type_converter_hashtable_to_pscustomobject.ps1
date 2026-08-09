# vybe-test: powershell/type_converters/type_converter_hashtable_to_pscustomobject
$hash = @{ Title = "PowerShell" }
$obj = [pscustomobject]$hash
if (-not ($obj -is [PSCustomObject]) -or $obj.Title -ne "PowerShell") {
    Write-Host "FAIL: hashtable to PSCustomObject converter failed"
    exit 1
}
Write-Host "PASS"
exit 0
