# vybe-test: powershell/switch/basic_switch
$value = 2
$result = switch ($value) { 1 { 'one' } 2 { 'two' } default { 'other' } }
if ($result -ne 'two') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
