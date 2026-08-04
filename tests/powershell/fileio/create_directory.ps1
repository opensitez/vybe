# vybe-test: powershell/fileio/create_directory
$tempDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.Guid]::NewGuid().ToString())
New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
$exists = Test-Path -Path $tempDir
Remove-Item -Path $tempDir -Force
if ($exists -ne $true) {
    Write-Host "FAIL: expected directory to exist"
    exit 1
}
Write-Host "PASS"
exit 0
