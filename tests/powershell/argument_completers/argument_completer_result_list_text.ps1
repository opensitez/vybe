# vybe-test: powershell/argument_completers/argument_completer_result_list_text
$cr = [System.Management.Automation.CompletionResult]::new("compl", "DisplayListText", [System.Management.Automation.CompletionResultType]::ParameterValue, "tip")
if ($cr.ListItemText -ne "DisplayListText") {
    Write-Host "FAIL: CompletionResult ListItemText expected 'DisplayListText', got '$($cr.ListItemText)'"
    exit 1
}
Write-Host "PASS"
exit 0
