# vybe-test: powershell/whitespace_and_line_rules/trailing_space_ignored
$left  = 4
$right = 6
$sum = $left + $right     

if ($sum -ne 10) {
    Write-Host "FAIL: trailing whitespace changed expression result: $sum"
    exit 1
}

Write-Host 'PASS'
exit 0
