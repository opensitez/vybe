# vybe-test: powershell/get_unique_cmdlet/get_unique_inputobject_parameter
# Get-Unique accepts collections directly through the -InputObject parameter
$res = Get-Unique -InputObject @(1, 1, 2)

if ($null -eq $res) {
    Write-Host "FAIL: Get-Unique -InputObject returned `$null"
    exit 1
}

Write-Host "PASS"
exit 0
