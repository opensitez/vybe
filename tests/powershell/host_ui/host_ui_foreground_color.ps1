# vybe-test: powershell/host_ui/host_ui_foreground_color
$color = $Host.UI.RawUI.ForegroundColor
if (-not $color) {
    Write-Host "FAIL: expected foreground color"
    exit 1
}
Write-Host 'PASS'
exit 0
