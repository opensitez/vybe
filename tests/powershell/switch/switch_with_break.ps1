# vybe-test: powershell/switch/switch_with_break
$value = 1
$result = switch ($value) { 1 { 'one'; break } 2 { 'two' } }
if ($result -ne 'one') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
