# vybe-test: powershell/culture_and_globalization_cmdlets/culture_pipeline_foreach_to_get_culture
# Culture names collected via ForEach-Object pipeline can each be resolved into CultureInfo
$queried = @("en-US", "fr-FR") | ForEach-Object { Get-Culture -Name $_ }

if ($queried.Count -ne 2) {
    Write-Host "FAIL: expected 2 cultures from pipeline query, got $($queried.Count)"
    exit 1
}

if ($queried[0].Name -ne "en-US" -or $queried[1].Name -ne "fr-FR") {
    Write-Host "FAIL: culture names mismatch: $($queried[0].Name), $($queried[1].Name)"
    exit 1
}

Write-Host "PASS"
exit 0
