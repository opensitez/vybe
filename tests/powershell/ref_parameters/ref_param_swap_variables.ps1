# vybe-test: powershell/ref_parameters/ref_param_swap_variables
function Swap-Vars([ref]$a, [ref]$b) {
    $temp = $a.Value
    $a.Value = $b.Value
    $b.Value = $temp
}
$first = "Alpha"
$second = "Omega"
Swap-Vars ([ref]$first) ([ref]$second)
if ($first -ne "Omega" -or $second -ne "Alpha") {
    Write-Host "FAIL: Swap-Vars expected first=Omega, second=Alpha"
    exit 1
}
Write-Host "PASS"
exit 0
