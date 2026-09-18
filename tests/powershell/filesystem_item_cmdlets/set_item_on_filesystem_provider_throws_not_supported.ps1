# vybe-test: powershell/filesystem_item_cmdlets/set_item_on_filesystem_provider_throws_not_supported
# Set-Item throws a NotSupportedException on the FileSystem provider because file content is modified via Set-Content
$tmp = [System.IO.Path]::GetTempFileName()
$threw = $false

try {
    Set-Item -Path $tmp -Value "illegal operation" -ErrorAction Stop
} catch {
    $threw = $true
} finally {
    if (Test-Path $tmp) {
        Remove-Item $tmp -Force
    }
}

if (-not $threw) {
    Write-Host "FAIL: Set-Item on FileSystem file did not throw"
    exit 1
}

Write-Host "PASS"
exit 0
