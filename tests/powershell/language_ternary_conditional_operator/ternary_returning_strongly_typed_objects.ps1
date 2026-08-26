# vybe-test: powershell/language_ternary_conditional_operator/ternary_returning_strongly_typed_objects
$flag = $true
$g1 = [guid]::Parse("11111111-1111-1111-1111-111111111111")
$g2 = [guid]::Parse("22222222-2222-2222-2222-222222222222")
$res = $flag ? $g1 : $g2
if ($res -ne $g1) {
    Write-Host "FAIL: Ternary returning strongly typed objects failed"
    exit 1
}
Write-Host "PASS"
exit 0
