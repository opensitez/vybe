# vybe-test: powershell/switch/array_switch
$values = 1,2,3
$result = switch ($values) { 2 { 'found' } default { 'notfound' } }
if ($result -ne 'found') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
