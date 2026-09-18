# vybe-test: powershell/compare_object_cmdlet/compare_object_syncwindow_zero_order_sensitivity
# Setting -SyncWindow 0 forces position-by-position comparison, rendering differently-ordered arrays different
$diff = Compare-Object @(1, 2) @(2, 1) -SyncWindow 0

if ($diff.Count -ne 4) {
    Write-Host "FAIL: expected 4 differences with -SyncWindow 0 for reversed array, got $($diff.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
