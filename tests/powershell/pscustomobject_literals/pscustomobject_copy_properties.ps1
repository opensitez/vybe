# vybe-test: powershell/pscustomobject_literals/pscustomobject_copy_properties
$source = [pscustomobject]@{ A = 1 }
$copy = [pscustomobject]@{ A = $source.A; B = 2 }
if ($copy.A -ne 1 -or $copy.B -ne 2) {
    Write-Host "FAIL: copy property expected A=1, B=2"
    exit 1
}
Write-Host "PASS"
exit 0
