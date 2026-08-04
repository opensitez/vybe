# vybe-test: powershell/fileio/delete_file
$tempFile = [System.IO.Path]::GetTempFileName()
Remove-Item -Path $tempFile
$exists = Test-Path -Path $tempFile
if ($exists -ne $false) {
    Write-Host "FAIL: expected file to be deleted"
    exit 1
}
Write-Host "PASS"
exit 0
