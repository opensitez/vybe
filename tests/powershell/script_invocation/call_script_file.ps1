# vybe-test: powershell/script_invocation/call_script_file
$scriptPath = [io.path]::Combine($PWD, 'calltemp.ps1')
'param($x) Write-Output ($x * 2)' | Out-File -FilePath $scriptPath -Encoding utf8
if ((& $scriptPath 3) -eq 6) { Remove-Item $scriptPath -ErrorAction SilentlyContinue; exit 0 }
Remove-Item $scriptPath -ErrorAction SilentlyContinue
exit 1
