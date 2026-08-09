# vybe-test: powershell/psscriptmethod_members/psscriptmethod_enumeration
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType ScriptMethod -Name "M1" -Value { 1 }
$methods = $obj.psobject.Members | Where-Object { $_.MemberType -eq "ScriptMethod" }
if ($methods.Count -ne 1 -or $methods[0].Name -ne "M1") {
    Write-Host "FAIL: ScriptMethod member enumeration expected M1"
    exit 1
}
Write-Host "PASS"
exit 0
