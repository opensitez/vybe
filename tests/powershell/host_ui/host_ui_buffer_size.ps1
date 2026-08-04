# vybe-test: powershell/host_ui/host_ui_buffer_size
$size = $Host.UI.RawUI.BufferSize
if (-not $size) {
    Write-Host "FAIL: expected buffer size"
    exit 1
}
Write-Host 'PASS'
exit 0
