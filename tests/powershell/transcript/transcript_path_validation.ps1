# vybe-test: powershell/transcript/transcript_path_validation
$path = Join-Path $PWD 'transcript-validate.txt'
Start-Transcript -Path $path -ErrorAction SilentlyContinue
Stop-Transcript -ErrorAction SilentlyContinue
if ($path -notlike '*transcript-validate.txt') {
    Write-Host "FAIL: expected valid path"
    exit 1
}
Remove-Item $path -ErrorAction SilentlyContinue
Write-Host 'PASS'
exit 0
