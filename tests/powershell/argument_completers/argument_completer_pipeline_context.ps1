# vybe-test: powershell/argument_completers/argument_completer_pipeline_context
$completer = {
    param($cmd, $param, $word, $ast, $bound)
    $word | ForEach-Object { "P:$_" }
}
$res = &$completer "W"
if ($res -ne "P:W") {
    Write-Host "FAIL: argument completer pipeline context expected P:W, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
