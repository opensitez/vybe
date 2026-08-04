# vybe-test: powershell/switch/no_match_default
$value = 0
$result = switch ($value) { 1 { 'one' } default { 'other' } }
if ($result -ne 'other') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
