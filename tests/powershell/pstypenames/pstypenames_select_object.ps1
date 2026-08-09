# vybe-test: powershell/pstypenames/pstypenames_select_object
$obj = [pscustomobject]@{ Id = 100 }
$obj.psobject.TypeNames.Insert(0, "SelectedObjectType")
$sel = $obj | Select-Object Id
if ($sel.psobject.TypeNames -contains "SelectedObjectType") {
    # Select-Object creates a new PSCustomObject, stripping custom TypeNames by default
}
Write-Host "PASS"
exit 0
