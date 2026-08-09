# vybe-test: powershell/argument_completers/argument_completer_scriptblock
$completerSb = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    return "SuggestionA", "SuggestionB"
}
$res = &$completerSb "Test-Cmd" "Param" "Sug" $null $null
if ($res[0] -ne "SuggestionA" -or $res[1] -ne "SuggestionB") {
    Write-Host "FAIL: argument completer scriptblock expected SuggestionA, SuggestionB"
    exit 1
}
Write-Host "PASS"
exit 0
