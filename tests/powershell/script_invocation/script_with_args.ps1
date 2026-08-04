# vybe-test: powershell/script_invocation/script_with_args
$scriptFile = [io.path]::Combine($PWD, 'args.ps1')
'param($a, $b) Write-Output "$a,$b"' | Out-File -FilePath $scriptFile -Encoding utf8
if ((& $scriptFile 1 2) -eq '1,2') { Remove-Item $scriptFile -ErrorAction SilentlyContinue; exit 0 }
Remove-Item $scriptFile -ErrorAction SilentlyContinue
exit 1
