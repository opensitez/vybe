# vybe-test: powershell/pstypenames/pstypenames_scriptblock_access
$obj = [pscustomobject]@{ Id = 5 }
$obj.psobject.TypeNames.Insert(0, "ScriptType")
$sb = { param($o) $o.psobject.TypeNames[0] }
$res = &$sb $obj
if ($res -ne "ScriptType") {
    Write-Host "FAIL: scriptblock TypeNames access expected ScriptType, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
