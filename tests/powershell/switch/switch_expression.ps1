# vybe-test: powershell/switch/switch_expression
$value = 2
$result = switch ($value) { 1 { 'one' } 2 { 'two' } }
if ($result -ne 'two') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
