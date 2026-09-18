# vybe-test: powershell/new_object_cmdlet/new_object_psobject_property_hashtable
# New-Object PSObject -Property @{...} creates an object with named NoteProperty members
$obj = New-Object PSObject -Property @{ Name = "Alice"; Age = 30 }

if ($obj.Name -ne "Alice" -or $obj.Age -ne 30) {
    Write-Host "FAIL: PSObject property mismatch, Name='$($obj.Name)', Age=$($obj.Age)"
    exit 1
}

Write-Host "PASS"
exit 0
