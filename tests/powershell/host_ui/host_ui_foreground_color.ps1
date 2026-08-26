# vybe-test: powershell/host_ui/host_ui_foreground_color
$x = 10
$x += 5
$x *= 2
if ($x -eq 30) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
