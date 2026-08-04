# vybe-test: powershell/transcript/transcript_file_created
$path = Join-Path $PWD 'transcript-file.txt'
Start-Transcript -Path $path -ErrorAction SilentlyContinue
Stop-Transcript -ErrorAction SilentlyContinue
if (-not (Test-Path $path)) {
    Write-Host "FAIL: expected file created"
    exit 1
}
Remove-Item $path -ErrorAction SilentlyContinue
Write-Host 'PASS'
exit 0
