# vybe-test: powershell/command_aliases/global_alias
New-Alias -Name globalg -Value Write-Host -Scope Global
if ((globalg 'hi') -eq 'hi') { exit 0 }
exit 1
