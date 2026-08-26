# vybe-test: powershell/collections_generic_list/exists_predicate
$list = [System.Collections.Generic.List[string]]::new([string[]]@("alpha", "beta", "gamma"))
$hasLong = $list.Exists([System.Predicate[string]]{ param($s) $s.Length -gt 4 })
if (-not $hasLong) {
    Write-Host "FAIL: Exists predicate failed"
    exit 1
}
Write-Host "PASS"
exit 0
