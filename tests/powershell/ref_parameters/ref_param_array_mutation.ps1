# vybe-test: powershell/ref_parameters/ref_param_array_mutation
function Push-Item([ref]$arrRef, $item) {
    $arrRef.Value += $item
}
$list = @(1, 2)
Push-Item ([ref]$list) 3
if ($list.Count -ne 3 -or $list[2] -ne 3) {
    Write-Host "FAIL: [ref] array push expected Count 3, item 3"
    exit 1
}
Write-Host "PASS"
exit 0
