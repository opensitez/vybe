# vybe-test: powershell/script_invocation/script_local_variable
$scriptFile = [io.path]::Combine($PWD, 'localvar.ps1')
'$a = 1; Write-Output $a' | Out-File -FilePath $scriptFile -Encoding utf8
if ((& $scriptFile) -eq 1) { Remove-Item $scriptFile -ErrorAction SilentlyContinue; exit 0 }
Remove-Item $scriptFile -ErrorAction SilentlyContinue
exit 1
