# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_iteration_in_foreach_loop
$obj = [pscustomobject]@{ X = 1 }
$collected = [System.Collections.Generic.List[string]]::new()
foreach ($tn in $obj.PSObject.TypeNames) {
    $collected.Add($tn)
}
if ($collected.Count -eq 0 -or -not ($collected -contains "System.Object")) {
    Write-Host "FAIL: PSTypeNames iteration failed"
    exit 1
}
Write-Host "PASS"
exit 0
