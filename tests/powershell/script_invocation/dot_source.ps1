# vybe-test: powershell/script_invocation/dot_source
$scriptFile = [io.path]::Combine($PWD, 'dottemp.ps1')
'param($x) Write-Output $x' | Out-File -FilePath $scriptFile -Encoding utf8
. $scriptFile 3
if ($x -eq 3) { Remove-Item $scriptFile -ErrorAction SilentlyContinue; exit 0 }
Remove-Item $scriptFile -ErrorAction SilentlyContinue
exit 1
