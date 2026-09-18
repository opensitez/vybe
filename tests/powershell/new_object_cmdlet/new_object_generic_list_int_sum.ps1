# vybe-test: powershell/new_object_cmdlet/new_object_generic_list_int_sum
# New-Object "System.Collections.Generic.List[int]" supports Add() and Measure-Object -Sum
$ints = New-Object "System.Collections.Generic.List[int]"
$ints.Add(10)
$ints.Add(20)
$ints.Add(30)

$sum = $ints | Measure-Object -Sum | Select-Object -ExpandProperty Sum

if ($sum -ne 60) {
    Write-Host "FAIL: expected sum 60, got $sum"
    exit 1
}

Write-Host "PASS"
exit 0
