# vybe-test: powershell/fileio/test_path_exists
$path = [System.IO.Path]::GetTempFileName()
if (-not (Test-Path $path)) {
    Write-Host "FAIL: file should exist"
    exit 1
}
Remove-Item $path
if (Test-Path $path) {
    Write-Host "FAIL: file should not exist after removal"
    exit 1
}
Write-Host "PASS"
exit 0
