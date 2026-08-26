# vybe-test: powershell/collections_generic_list/contains_check
$list = [System.Collections.Generic.List[string]]::new([string[]]@("cat", "dog"))
if (-not $list.Contains("dog") -or $list.Contains("fish")) {
    Write-Host "FAIL: Contains check failed"
    exit 1
}
Write-Host "PASS"
exit 0
