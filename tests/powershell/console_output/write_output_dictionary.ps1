# vybe-test: powershell/console_output/write_output_dictionary
$hash = @{ A = 1 }
$result = @(Write-Output $hash)
if ($result[0].A -ne 1) {
    Write-Host "FAIL: expected A=1"
    exit 1
}
Write-Host 'PASS'
exit 0
