# vybe-test: powershell/pscustomobject_literals/pscustomobject_basic
$obj = [pscustomobject]@{ Id = 42; Name = "Core" }
if ($obj.Id -ne 42) {
    Write-Host "FAIL: Id expected 42, got $($obj.Id)"
    exit 1
}
if ($obj.Name -ne "Core") {
    Write-Host "FAIL: Name expected Core, got $($obj.Name)"
    exit 1
}
Write-Host "PASS"
exit 0
