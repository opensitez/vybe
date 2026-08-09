# vybe-test: powershell/psscriptmethod_members/psscriptmethod_null_args
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType ScriptMethod -Name "CheckNull" -Value { param($param) if ($param -eq $null) { "IS_NULL" } }
$res = $obj.CheckNull($null)
if ($res -ne "IS_NULL") {
    Write-Host "FAIL: PSScriptMethod null parameter passing expected IS_NULL"
    exit 1
}
Write-Host "PASS"
exit 0
