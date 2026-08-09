# vybe-test: powershell/argument_completers/argument_completer_parameter_name
$completer = {
    param($cmd, $param, $word, $ast, $bound)
    return "PARAM:$param"
}
$res = &$completer "Cmd" "TargetParam" "" $null $null
if ($res -ne "PARAM:TargetParam") {
    Write-Host "FAIL: argument completer parameterName expected PARAM:TargetParam, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
