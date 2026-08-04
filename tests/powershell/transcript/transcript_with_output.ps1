# vybe-test: powershell/transcript/transcript_with_output
$path = Join-Path $PWD 'transcript-output.txt'
Start-Transcript -Path $path -ErrorAction SilentlyContinue
Write-Host 'hello'
Stop-Transcript -ErrorAction SilentlyContinue
Remove-Item $path -ErrorAction SilentlyContinue
Write-Host 'PASS'
exit 0
