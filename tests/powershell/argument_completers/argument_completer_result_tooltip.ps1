# vybe-test: powershell/argument_completers/argument_completer_result_tooltip
$cr = [System.Management.Automation.CompletionResult]::new("text", "listText", [System.Management.Automation.CompletionResultType]::ParameterValue, "Custom ToolTip")
if ($cr.ToolTip -ne "Custom ToolTip") {
    Write-Host "FAIL: CompletionResult ToolTip expected 'Custom ToolTip', got '$($cr.ToolTip)'"
    exit 1
}
Write-Host "PASS"
exit 0
