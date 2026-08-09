# vybe-test: powershell/pipeline_chaining/chain_scriptblock_eval
$sb = { ($args[0] -gt 5) && "Pass" }
$res = &$sb 10
if ($res -ne "Pass") {
    Write-Host "FAIL: scriptblock pipeline chaining expected 'Pass', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
