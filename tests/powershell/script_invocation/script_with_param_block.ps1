# vybe-test: powershell/script_invocation/script_with_param_block.ps1
$scriptFile = [io.path]::Combine($PWD, 'paramblock.ps1')
'param($x) Write-Output $x' | Out-File -FilePath $scriptFile -Encoding utf8
if ((& $scriptFile 9) -eq 9) { Remove-Item $scriptFile -ErrorAction SilentlyContinue; exit 0 }
Remove-Item $scriptFile -ErrorAction SilentlyContinue
exit 1
