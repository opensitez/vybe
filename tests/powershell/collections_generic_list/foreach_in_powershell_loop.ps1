# vybe-test: powershell/collections_generic_list/foreach_in_powershell_loop
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2, 3, 4))
$sum = 0
foreach ($item in $list) { $sum += $item }
if ($sum -ne 10) {
    Write-Host "FAIL: foreach loop on generic List failed, expected 10, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
