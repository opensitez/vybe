# vybe-test: powershell/switch/fallthrough_switch
$count = 0
switch (1) { 1 { $count++; continue } 2 { $count++ } }
if ($count -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
