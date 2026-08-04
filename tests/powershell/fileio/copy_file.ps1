# vybe-test: powershell/fileio/copy_file
$tempFile = [System.IO.Path]::GetTempFileName()
$destFile = $tempFile + ".copy"
Set-Content -Path $tempFile -Value "test content"
Copy-Item -Path $tempFile -Destination $destFile
$exists = Test-Path -Path $destFile
Remove-Item -Path $tempFile, $destFile -Force
if ($exists -ne $true) {
    Write-Host "FAIL: expected copied file to exist"
    exit 1
}
Write-Host "PASS"
exit 0
