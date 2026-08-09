# vybe-test: powershell/scriptblock_closures/closure_module_context
$m = New-Module -ScriptBlock {
    $secret = "ModSecret"
    function Get-SecretSb {
        return { $secret }.GetClosure()
    }
}
$sb = & $m { Get-SecretSb }
$res = &$sb
if ($res -ne "ModSecret") {
    Write-Host "FAIL: module context closure expected ModSecret, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
