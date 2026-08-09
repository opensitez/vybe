# vybe-test: powershell/argument_completers/argument_completer_empty_word
$completer = {
    param($cmd, $param, $word, $ast, $bound)
    if ($word -eq "") { @("All1", "All2") } else { $word }
}
$res = @(&$completer "" "" "" $null $null)
if ($res.Count -ne 2 -or $res[0] -ne "All1") {
    Write-Host "FAIL: empty word completion expected All1, All2"
    exit 1
}
Write-Host "PASS"
exit 0
