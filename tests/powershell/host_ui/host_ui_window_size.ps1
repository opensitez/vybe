# vybe-test: powershell/host_ui/host_ui_window_size
$size = $Host.UI.RawUI.WindowSize
if (-not $size.Width) {
    Write-Host "FAIL: expected window size width"
    exit 1
}
Write-Host 'PASS'
exit 0
