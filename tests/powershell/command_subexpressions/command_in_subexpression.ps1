# vybe-test: powershell/command_subexpressions/command_in_subexpression
$res = $( $a = 5; $b = 10; $a + $b )
if ($res -ne 15) {
    Write-Host "FAIL: Command in subexpression failed"
    exit 1
}
Write-Host "PASS"
exit 0
