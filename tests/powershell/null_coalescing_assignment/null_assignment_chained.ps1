# vybe-test: powershell/null_coalescing_assignment/null_assignment_chained
$a = $null
$b = $null
$c = "Primary"
$a ??= $b ??= $c
if ($a -ne "Primary" -or $b -ne "Primary") {
    Write-Host "FAIL: chained ??= expected a=Primary, b=Primary"
    exit 1
}
Write-Host "PASS"
exit 0
