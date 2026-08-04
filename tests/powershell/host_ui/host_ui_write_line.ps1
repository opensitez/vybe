# vybe-test: powershell/host_ui/host_ui_write_line
$Host.UI.Write('x')
if (-not $Host.UI) {
    Write-Host "FAIL: expected host UI"
    exit 1
}
Write-Host 'PASS'
exit 0
