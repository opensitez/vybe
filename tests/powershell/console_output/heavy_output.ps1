# vybe-test: powershell/console_output/heavy_output
$result = @(1..5 | ForEach-Object { Write-Output $_ })
if ($result.Count -ne 5) {
    Write-Host "FAIL: expected 5 outputs"
    exit 1
}
Write-Host 'PASS'
exit 0
