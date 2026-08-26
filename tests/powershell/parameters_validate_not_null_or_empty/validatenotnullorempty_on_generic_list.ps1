# vybe-test: powershell/parameters_validate_not_null_or_empty/validatenotnullorempty_on_generic_list
function Test-ListNotNullOrEmpty {
    param([ValidateNotNullOrEmpty()][System.Collections.Generic.List[int]]$List)
    return $List.Count
}
$l = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2, 3))
$res = Test-ListNotNullOrEmpty -List $l
if ($res -ne 3) {
    Write-Host "FAIL: Generic List ValidateNotNullOrEmpty failed"
    exit 1
}
Write-Host "PASS"
exit 0
