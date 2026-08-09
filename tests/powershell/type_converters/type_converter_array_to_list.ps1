# vybe-test: powershell/type_converters/type_converter_array_to_list
$arr = @(1, 2, 3)
$list = [System.Collections.Generic.List[int]]$arr
if (-not ($list -is [System.Collections.Generic.List[int]]) -or $list.Count -ne 3) {
    Write-Host "FAIL: array to List[int] converter failed"
    exit 1
}
Write-Host "PASS"
exit 0
