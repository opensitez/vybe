# vybe-test: powershell/break_continue/continue_do
$result = 0
do { $result++; if ($result -eq 2) { continue }; if ($result -eq 3) { break } } while ($true)
if ($result -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
