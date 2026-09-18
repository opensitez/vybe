# vybe-test: powershell/new_object_cmdlet/new_object_hashtable_type
# New-Object System.Collections.Hashtable creates an empty Hashtable
$ht = New-Object System.Collections.Hashtable

if ($ht.GetType().Name -ne "Hashtable") {
    Write-Host "FAIL: expected Hashtable type, got: $($ht.GetType().Name)"
    exit 1
}

$ht["key"] = "value"
if ($ht["key"] -ne "value") {
    Write-Host "FAIL: Hashtable key assignment failed"
    exit 1
}

Write-Host "PASS"
exit 0
