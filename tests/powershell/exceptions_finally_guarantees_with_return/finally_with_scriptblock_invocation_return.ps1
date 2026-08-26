# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_with_scriptblock_invocation_return
$ran = $false
$sb = {
    try {
        return "SbResult"
    } finally {
        $script:ran = $true
    }
}
$res = & $sb
if ($res -ne "SbResult" -or -not $ran) {
    Write-Host "FAIL: Finally in scriptblock invocation failed"
    exit 1
}
Write-Host "PASS"
exit 0
