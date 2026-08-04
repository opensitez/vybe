# vybe-test: powershell/ranges/range_subexpression
$range = @(1..3)
if ($range.Count -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
