# vybe-test: powershell/parameters_validate_count/validatecount_with_object_array
function Process-Objects {
    param([ValidateCount(1, 3)][object[]]$Objs)
    return $Objs.Length
}
$res = Process-Objects -Objs 1, "two", 3.0
if ($res -ne 3) {
    Write-Host "FAIL: ValidateCount object array failed"
    exit 1
}
Write-Host "PASS"
exit 0
