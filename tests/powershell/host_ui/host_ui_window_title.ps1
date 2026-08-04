# vybe-test: powershell/host_ui/host_ui_window_title
$title = $Host.UI.RawUI.WindowTitle
if (-not $title) {
    Write-Host "FAIL: expected window title"
    exit 1
}
Write-Host 'PASS'
exit 0
