# vybe-test: powershell/host_ui/host_ui_cursor_position
$pos = $Host.UI.RawUI.CursorPosition
if (-not $pos) {
    Write-Host "FAIL: expected cursor position"
    exit 1
}
Write-Host 'PASS'
exit 0
