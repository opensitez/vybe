# vybe-test: powershell/script_invocation/nested_script_invocation
$scriptFile = [io.path]::Combine($PWD, 'nested.ps1')
'Write-Output 2' | Out-File -FilePath $scriptFile -Encoding utf8
if ((& $scriptFile | ForEach-Object { $_ }) -eq 2) { Remove-Item $scriptFile -ErrorAction SilentlyContinue; exit 0 }
Remove-Item $scriptFile -ErrorAction SilentlyContinue
exit 1
