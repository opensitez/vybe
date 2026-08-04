# vybe-test: powershell/objects/pscustomobject_create
$obj = [PSCustomObject]@{ Name = "John"; Age = 30 }
if ($obj.Name -ne "John") {
    Write-Host "FAIL: expected 'John', got '$($obj.Name)'"
    exit 1
}
if ($obj.Age -ne 30) {
    Write-Host "FAIL: expected 30, got $($obj.Age)"
    exit 1
}
Write-Host "PASS"
exit 0
