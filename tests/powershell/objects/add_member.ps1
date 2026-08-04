# vybe-test: powershell/objects/add_member
$obj = [PSCustomObject]@{ Name = "John" }
$obj | Add-Member -MemberType NoteProperty -Name Age -Value 30
if ($obj.Age -ne 30) {
    Write-Host "FAIL: expected Age to be 30, got $($obj.Age)"
    exit 1
}
Write-Host "PASS"
exit 0
