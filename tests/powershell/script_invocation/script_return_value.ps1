# vybe-test: powershell/script_invocation/script_return_value
$scriptFile = [io.path]::Combine($PWD, 'returnvalue.ps1')
'Write-Output 8' | Out-File -FilePath $scriptFile -Encoding utf8
if ((& $scriptFile) -eq 8) { Remove-Item $scriptFile -ErrorAction SilentlyContinue; exit 0 }
Remove-Item $scriptFile -ErrorAction SilentlyContinue
exit 1
