# vybe-test: powershell/ranges/empty_range
$range = 3..1
if ($range.Length -ne 3) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'PASS'
exit 0
