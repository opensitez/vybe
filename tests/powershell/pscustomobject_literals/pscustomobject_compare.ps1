# vybe-test: powershell/pscustomobject_literals/pscustomobject_compare
$o1 = [pscustomobject]@{ A = 1 }
$o2 = [pscustomobject]@{ A = 1 }
if ($o1 -eq $o2) {
    Write-Host "FAIL: different PSCustomObject reference instances should not be reference equal"
    exit 1
}
Write-Host "PASS"
exit 0
