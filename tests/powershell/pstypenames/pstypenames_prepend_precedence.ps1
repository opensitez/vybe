# vybe-test: powershell/pstypenames/pstypenames_prepend_precedence
$obj = [pscustomobject]@{ Val = 1 }
$obj.PSTypeNames.Insert(0, "SubKind")
$obj.PSTypeNames.Insert(0, "TopKind")
if ($obj.PSTypeNames[0] -ne "TopKind") {
    Write-Host "FAIL: prepended TypeName expected TopKind at index 0, got $($obj.PSTypeNames[0])"
    exit 1
}
Write-Host "PASS"
exit 0
