# vybe-test: powershell/invoke_expression_cmdlet/iex_alias_resolves_to_invoke_expression
# The 'iex' alias must resolve to the full cmdlet name 'Invoke-Expression'
$aliasTarget = (Get-Alias iex).ResolvedCommandName

if ($aliasTarget -ne "Invoke-Expression") {
    Write-Host "FAIL: 'iex' alias resolved to '$aliasTarget', expected 'Invoke-Expression'"
    exit 1
}

Write-Host "PASS"
exit 0
