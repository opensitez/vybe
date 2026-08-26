# vybe-test: powershell/language_ternary_conditional_operator/ternary_in_pipeline_foreach_object
$nums = @(1, 2, 3, 4)
$labels = @($nums | ForEach-Object { ($_ % 2 -eq 0) ? "Even" : "Odd" })
if ($labels[0] -ne "Odd" -or $labels[1] -ne "Even" -or $labels[3] -ne "Even") {
    Write-Host "FAIL: Ternary in pipeline ForEach-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
