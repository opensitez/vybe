# vybe-test: powershell/psdrive_management_cmdlets/psdrive_remove_by_literalname_parameter
# Remove-PSDrive supports the -LiteralName parameter to bypass wildcard resolution
$drive = New-PSDrive -Name LitParamCheckDrive -PSProvider FileSystem -Root $pwd.Path

Remove-PSDrive -LiteralName LitParamCheckDrive

$after = Get-PSDrive -Name LitParamCheckDrive -ErrorAction SilentlyContinue

if ($null -ne $after) {
    Write-Host "FAIL: drive was not removed via -LiteralName parameter"
    exit 1
}

Write-Host "PASS"
exit 0
