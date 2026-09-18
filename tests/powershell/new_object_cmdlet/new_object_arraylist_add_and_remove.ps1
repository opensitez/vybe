# vybe-test: powershell/new_object_cmdlet/new_object_arraylist_add_and_remove
# New-Object System.Collections.ArrayList supports dynamic Add() and Remove() operations
$al = New-Object System.Collections.ArrayList
$al.Add("alpha") | Out-Null
$al.Add("beta") | Out-Null
$al.Add("gamma") | Out-Null
$al.Remove("beta")

if ($al.Count -ne 2) {
    Write-Host "FAIL: expected Count 2 after Remove, got $($al.Count)"
    exit 1
}

if ($al -contains "beta") {
    Write-Host "FAIL: 'beta' should have been removed"
    exit 1
}

Write-Host "PASS"
exit 0
