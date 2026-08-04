# vybe-test: powershell/ranges/foreach_range
$count = 0
foreach ($i in 1..3) { $count++ }
if ($count -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
