# vybe-test: powershell/argument_completers/argument_completer_completion_result_type
$cr = [System.Management.Automation.CompletionResult]::new("val1", "val1", [System.Management.Automation.CompletionResultType]::ParameterValue, "TooltipText")
if ($cr.CompletionText -ne "val1" -or $cr.ToolTip -ne "TooltipText") {
    Write-Host "FAIL: CompletionResult constructor property assignment failed"
    exit 1
}
Write-Host "PASS"
exit 0
