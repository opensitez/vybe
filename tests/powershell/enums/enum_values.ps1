# vybe-test: powershell/enums/enum_values
enum Status {
    Pending
    Active
    Completed
}

$status = [Status]::Active
if ($status -ne "Active") {
    Write-Host "FAIL: expected 'Active', got '$status'"
    exit 1
}
Write-Host "PASS"
exit 0
