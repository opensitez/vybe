# vybe-test: powershell/using_variable_scope/using_variable_in_scriptblock
$msg = "ScriptBlockMsg"
$sb = [scriptblock]::Create("`$using:msg")
$res = &$sb
if ($res -ne "ScriptBlockMsg") {
    Write-Host "FAIL: dynamic scriptblock creation with \$using:msg expected ScriptBlockMsg"
    exit 1
}
Write-Host "PASS"
exit 0
