# vybe-test: powershell/null_coalescing/nested_null_coalescing
$a = $null
$b = $null
$result = $a ?? $b ?? 'final'
if ($result -ne 'final') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
