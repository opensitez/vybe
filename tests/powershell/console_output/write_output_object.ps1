# vybe-test: powershell/console_output/write_output_object
$obj = [PSCustomObject]@{ Name = 'x' }
$r = @(Write-Output $obj)
if ($r[0].Name -ne 'x') {
    Write-Host "FAIL: expected name x"
    exit 1
}
Write-Host 'PASS'
exit 0
