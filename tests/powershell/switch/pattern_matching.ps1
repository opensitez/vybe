# vybe-test: powershell/switch/pattern_matching
$value = 'hello'
$result = switch ($value) { 'h*' { 'match' } default { 'miss' } }
if ($result -ne 'match') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
