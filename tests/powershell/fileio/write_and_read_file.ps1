# vybe-test: powershell/fileio/write_and_read_file
$path = [System.IO.Path]::GetTempFileName()
Set-Content -Path $path -Value "hello from powershell"
$content = Get-Content -Path $path
if ($content -ne "hello from powershell") {
    Write-Host "FAIL: expected 'hello from powershell', got '$content'"
    Remove-Item $path -ErrorAction SilentlyContinue
    exit 1
}
Remove-Item $path
Write-Host "PASS"
exit 0
