# vybe-test: powershell/indexing/hashtable_property_indexing
$obj = [pscustomobject]@{ Foo = 'bar' }
if ($obj['Foo'] -ne 'bar') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
