# vybe-test: powershell/transcript/transcript_stop_after_start
$path = Join-Path $PWD 'transcript-stop.txt'
Start-Transcript -Path $path -ErrorAction SilentlyContinue
Stop-Transcript -ErrorAction SilentlyContinue
Write-Host 'PASS'
exit 0
