# vybe-test: powershell/pstypenames/pstypenames_subexpression
$obj = [pscustomobject]@{ Val = 1 }
$obj.psobject.TypeNames.Insert(0, "SubExprType")
$msg = "Type: $( $obj.psobject.TypeNames[0] )"
if ($msg -ne "Type: SubExprType") {
    Write-Host "FAIL: TypeNames in subexpression expected 'Type: SubExprType', got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
