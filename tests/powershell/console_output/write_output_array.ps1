# vybe-test: powershell/console_output/write_output_array
$result = @(Write-Output 1,2,3)
if ($result.Count -ne 3) {
    Write-Host "FAIL: expected 3 items"
    exit 1
}
Write-Host 'PASS'
exit 0
