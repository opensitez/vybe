# vybe-test: powershell/readonly_variables/readonly_variable_scriptblock_access
New-Variable -Name "RO_SB_DATA" -Value "SB_Read" -Option ReadOnly
$sb = { $RO_SB_DATA }
$res = &$sb
if ($res -ne "SB_Read") {
    Write-Host "FAIL: scriptblock ReadOnly variable access expected SB_Read, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
