# vybe-test: powershell/indexing/multiple_dimensions
$matrix = @(,@(1,2),@(3,4))
if ($matrix[1][0] -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
