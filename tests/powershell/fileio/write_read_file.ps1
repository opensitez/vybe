# vybe-test: powershell/fileio/write_read_file
$tempFile = [System.IO.Path]::GetTempFileName()
$content = "Hello, PowerShell!"
Set-Content -Path $tempFile -Value $content
$read = Get-Content -Path $tempFile
Remove-Item -Path $tempFile
if ($read -ne $content) {
    Write-Host "FAIL: expected '$content', got '$read'"
    exit 1
}
Write-Host "PASS"
exit 0
