# vybe-test: powershell/argument_completers/argument_completer_pipeline_context
$completer = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    [System.Management.Automation.CompletionResult]::new("CompletedWord", "CompletedWord", [System.Management.Automation.CompletionResultType]::ParameterValue, "Help")
}
$res = & $completer "cmd" "param" "comp" $null $null
if ($res.CompletionText -ne "CompletedWord") {
    Write-Host "FAIL: Argument completer failed"
    exit 1
}
Write-Host "PASS"
exit 0
