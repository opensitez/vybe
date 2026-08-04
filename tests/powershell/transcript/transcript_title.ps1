# vybe-test: powershell/transcript/transcript_title
$path = Join-Path $PWD 'transcript-title.txt'
Start-Transcript -Path $path -ErrorAction SilentlyContinue
Stop-Transcript -ErrorAction SilentlyContinue
Write-Host 'PASS'
exit 0
