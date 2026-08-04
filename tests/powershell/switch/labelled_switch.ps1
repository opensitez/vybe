# vybe-test: powershell/switch/labelled_switch
$value = 1
switch ($value) { 1 { 'one' } 2 { 'two' } }
Write-Host 'PASS'
exit 0
