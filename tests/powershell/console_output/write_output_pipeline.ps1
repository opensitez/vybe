# vybe-test: powershell/console_output/write_output_pipeline
$result = Invoke-Expression 'Write-Output 5 | Measure-Object | Select-Object -ExpandProperty Count'
if ($result -ne 1) {
    Write-Host "FAIL: expected count 1"
    exit 1
}
Write-Host 'PASS'
exit 0
