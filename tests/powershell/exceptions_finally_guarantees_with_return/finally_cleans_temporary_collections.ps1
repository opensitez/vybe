# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_cleans_temporary_collections
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2, 3))
function Clear-ListTemp([System.Collections.Generic.List[int]]$l) {
    try {
        return $l.Count
    } finally {
        $l.Clear()
    }
}
$count = Clear-ListTemp $list
if ($count -ne 3 -or $list.Count -ne 0) {
    Write-Host "FAIL: Collection clean in finally failed"
    exit 1
}
Write-Host "PASS"
exit 0
