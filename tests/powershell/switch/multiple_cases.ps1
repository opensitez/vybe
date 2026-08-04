# vybe-test: powershell/switch/multiple_cases
$value = 3
$result = switch ($value) { 1 { 'one' } 2 { 'two' } 3 { 'three' } }
if ($result -ne 'three') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
