# vybe-test: powershell/ref_parameters/ref_param_return_and_ref_combination
function Compute-And-Set([ref]$statusRef, [int]$a, [int]$b) {
    if ($a + $b -gt 100) {
        $statusRef.Value = "Overflow"
    } else {
        $statusRef.Value = "Normal"
    }
    return ($a + $b)
}
$status = ""
$total = Compute-And-Set ([ref]$status) 60 50
if ($total -ne 110 -or $status -ne "Overflow") {
    Write-Host "FAIL: return and ref combination expected total=110, status=Overflow"
    exit 1
}
Write-Host "PASS"
exit 0
