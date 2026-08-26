# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_in_pipeline_filtering_by_typename
$o1 = [pscustomobject]@{ Id = 1 }
$o1.PSObject.TypeNames.Insert(0, "TypeA")
$o2 = [pscustomobject]@{ Id = 2 }
$o2.PSObject.TypeNames.Insert(0, "TypeB")
$items = @($o1, $o2)
$filtered = @($items | Where-Object { $_.PSObject.TypeNames -contains "TypeA" })
if ($filtered.Length -ne 1 -or $filtered[0].Id -ne 1) {
    Write-Host "FAIL: Filtering pipeline by PSTypeNames failed"
    exit 1
}
Write-Host "PASS"
exit 0
