# vybe-test: powershell/host_ui/host_ui_read_line
if (-not $Host.UI.RawUI) {
    Write-Host "FAIL: expected RawUI"
    exit 1
}
Write-Host 'PASS'
exit 0
