# vybe-test: powershell/argument_completers/argument_completer_multiple_suggestions
$completer = {
    param($c, $p, $w, $a, $b)
    1..5 | ForEach-Object { [System.Management.Automation.CompletionResult]::new("Opt$_", "Opt$_", [System.Management.Automation.CompletionResultType]::ParameterValue, "Opt$_") }
}
$results = @(&$completer "" "" "" $null $null)
if ($results.Count -ne 5 -or $results[4].CompletionText -ne "Opt5") {
    Write-Host "FAIL: multiple CompletionResult suggestions count expected 5, last Opt5"
    exit 1
}
Write-Host "PASS"
exit 0
