# vybe-test: powershell/host_ui/host_ui_key_available
if (-not $Host.UI.RawUI.KeyAvailable -and $Host.UI.RawUI.KeyAvailable -ne $false) {
    Write-Host "FAIL: expected key available property"
    exit 1
}
Write-Host 'PASS'
exit 0
