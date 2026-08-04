# vybe-test: powershell/fileio/file_exists
$tempFile = [System.IO.Path]::GetTempFileName()
$exists = Test-Path -Path $tempFile
Remove-Item -Path $tempFile
if ($exists -ne $true) {
    Write-Host "FAIL: expected file to exist"
    exit 1
}
Write-Host "PASS"
exit 0
