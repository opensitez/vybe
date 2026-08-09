# vybe-test: powershell/psscriptmethod_members/psscriptmethod_type_check
$obj = [pscustomobject]@{}
$sm = [System.Management.Automation.PSScriptMethod]::new("Run", { 100 })
$obj.psobject.Members.Add($sm)
$m = $obj.psobject.Members["Run"]
if (-not ($m -is [System.Management.Automation.PSScriptMethod])) {
    Write-Host "FAIL: PSScriptMethod member type check failed"
    exit 1
}
Write-Host "PASS"
exit 0
