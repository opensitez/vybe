# vybe-test: powershell/backtick_continuation/basic_backtick
$x = 1 + `
2
if ($x -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
