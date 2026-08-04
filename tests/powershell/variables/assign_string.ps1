# vybe-test: powershell/variables/assign_string
$name = "PowerShell"
if ($name -ne "PowerShell") {
    Write-Host "FAIL: expected PowerShell, got $name"
    exit 1
}
Write-Host "PASS"
exit 0
