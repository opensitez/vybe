# vybe-test: powershell/host_ui/host_ui_background_color
$color = $Host.UI.RawUI.BackgroundColor
if (-not $color) {
    Write-Host "FAIL: expected background color"
    exit 1
}
Write-Host 'PASS'
exit 0
