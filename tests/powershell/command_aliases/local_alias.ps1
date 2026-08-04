# vybe-test: powershell/command_aliases/local_alias
function Test-Alias { Set-Alias localg Write-Host -Scope Local; return (localg 'x') }
if ((Test-Alias) -eq 'x') { exit 0 }
exit 1
