# vybe-test: powershell/scriptblock_closures/closure_reference_type_mutation
$dict = @{ Count = 0 }
$sb = { $dict.Count++ }.GetNewClosure()
&$sb
&$sb
if ($dict.Count -ne 2) {
    Write-Host "FAIL: reference type property mutation inside closure expected Count=2, got $($dict.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
