# vybe-test: powershell/transcript/start_transcript
$path = Join-Path $PWD 'transcript-test.txt'
Start-Transcript -Path $path -ErrorAction SilentlyContinue
Stop-Transcript -ErrorAction SilentlyContinue
if (-not (Test-Path $path)) {
    Write-Host "FAIL: expected transcript file"
    exit 1
}
Remove-Item $path -ErrorAction SilentlyContinue
Write-Host 'PASS'
exit 0
