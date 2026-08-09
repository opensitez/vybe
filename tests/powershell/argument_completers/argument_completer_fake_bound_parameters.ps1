# vybe-test: powershell/argument_completers/argument_completer_fake_bound_parameters
$boundParams = @{ Mode = "Verbose" }
$completer = {
    param($cmd, $param, $word, $ast, $bound)
    return $bound["Mode"]
}
$res = &$completer "Cmd" "P" "" $null $boundParams
if ($res -ne "Verbose") {
    Write-Host "FAIL: fakeBoundParameters expected Verbose, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
