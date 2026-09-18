# vybe-test: powershell/new_object_cmdlet/new_object_generic_list_string_construction
# New-Object with a generic type name creates a strongly-typed List[string]
$list = New-Object "System.Collections.Generic.List[string]"
$list.Add("alpha")
$list.Add("beta")

if ($list.Count -ne 2) {
    Write-Host "FAIL: expected Count 2, got $($list.Count)"
    exit 1
}

if ($list[0] -ne "alpha" -or $list[1] -ne "beta") {
    Write-Host "FAIL: element mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
