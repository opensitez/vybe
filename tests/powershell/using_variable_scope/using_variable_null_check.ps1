# vybe-test: powershell/using_variable_scope/using_variable_null_check
$emptyVal = $null
$sb = { if ($using:emptyVal -eq $null) { "IS_NULL" } }
$res = &$sb
if ($res -ne "IS_NULL") {
    Write-Host "FAIL: null using variable expected IS_NULL, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
