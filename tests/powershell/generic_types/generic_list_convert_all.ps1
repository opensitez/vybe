# vybe-test: powershell/generic_types/generic_list_convert_all
$list = [System.Collections.Generic.List[int]]::new()
$list.AddRange([int[]]@(1, 2, 3))
$strList = $list.ConvertAll([Converter[int, string]]{ param($i) "N$i" })
if ($strList[0] -ne "N1" -or $strList[2] -ne "N3") {
    Write-Host "FAIL: ConvertAll expected N1, N2, N3"
    exit 1
}
Write-Host "PASS"
exit 0
