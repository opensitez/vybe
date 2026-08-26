# vybe-test: powershell/dynamic_assembly_type_resolution/resolve_collection_accelerators
$tList = [type]"hashtable"
$tPso = [type]"pscustomobject"
if ($tList -ne [hashtable] -or $tPso -ne [pscustomobject]) {
    Write-Host "FAIL: Collection type accelerator resolution failed"
    exit 1
}
Write-Host "PASS"
exit 0
