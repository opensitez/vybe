# vybe-test: powershell/transcript/transcript_multiple_sessions
$path = Join-Path $PWD 'transcript-multi-sessions.txt'
Start-Transcript -Path $path -ErrorAction SilentlyContinue
Stop-Transcript -ErrorAction SilentlyContinue
Start-Transcript -Path $path -ErrorAction SilentlyContinue
Stop-Transcript -ErrorAction SilentlyContinue
Remove-Item $path -ErrorAction SilentlyContinue
Write-Host 'PASS'
exit 0
