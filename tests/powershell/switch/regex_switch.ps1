# vybe-test: powershell/switch/regex_switch
$value = 'abc123'
$result = switch -Regex ($value) { '\d+' { 'digits' } default { 'none' } }
if ($result -ne 'digits') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
