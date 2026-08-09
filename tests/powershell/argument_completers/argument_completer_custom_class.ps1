# vybe-test: powershell/argument_completers/argument_completer_custom_class
class CustomCompleter : System.Management.Automation.IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[System.Management.Automation.CompletionResult]] CompleteArgument(
        [string]$commandName,
        [string]$parameterName,
        [string]$wordToComplete,
        [System.Management.Automation.Language.CommandAst]$commandAst,
        [System.Collections.IDictionary]$fakeBoundParameters
    ) {
        $list = [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()
        $list.Add([System.Management.Automation.CompletionResult]::new("ClassOpt", "ClassOpt", [System.Management.Automation.CompletionResultType]::ParameterValue, "ClassOpt"))
        return $list
    }
}
$comp = [CustomCompleter]::new()
$res = @($comp.CompleteArgument("", "", "", $null, $null))
if ($res[0].CompletionText -ne "ClassOpt") {
    Write-Host "FAIL: IArgumentCompleter class implementation expected ClassOpt, got $($res[0].CompletionText)"
    exit 1
}
Write-Host "PASS"
exit 0
