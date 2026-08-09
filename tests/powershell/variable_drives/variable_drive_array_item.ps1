# vybe-test: powershell/variable_drives/variable_drive_array_item
$variable:items = @(1, 2, 3)
if ($items.Length -ne 3 -or $items[2] -ne 3) {
    Write-Host "FAIL: \$variable: items array expected 1, 2, 3"
    exit 1
}
Write-Host "PASS"
exit 0
