# vybe-test: powershell/subexpressions/multi_line_subexpression
$value = $(
    $a = 1
    $a + 1
)
if ($value -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
