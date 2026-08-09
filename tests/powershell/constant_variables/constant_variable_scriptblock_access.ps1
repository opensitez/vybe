# vybe-test: powershell/constant_variables/constant_variable_scriptblock_access
New-Variable -Name "SB_CONST" -Value "SB_Data" -Option Constant
$sb = { $SB_CONST }
$res = &$sb
if ($res -ne "SB_Data") {
    Write-Host "FAIL: scriptblock Constant variable access expected SB_Data, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
