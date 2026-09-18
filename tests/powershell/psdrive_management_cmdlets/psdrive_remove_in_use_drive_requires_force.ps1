# vybe-test: powershell/psdrive_management_cmdlets/psdrive_remove_in_use_drive_requires_force
# In PowerShell, removing a PSDrive when the current location is inside that drive throws unless -Force is used
$origLoc = $pwd.Path
$drive = New-PSDrive -Name InUseTestDrive -PSProvider FileSystem -Root $pwd.Path

Set-Location InUseTestDrive:

$threwWithoutForce = $false
try {
    Remove-PSDrive -Name InUseTestDrive -ErrorAction Stop
} catch {
    $threwWithoutForce = $true
}

Set-Location $origLoc
Remove-PSDrive -Name InUseTestDrive -Force

if (-not $threwWithoutForce) {
    Write-Host "FAIL: removing in-use drive without -Force did not throw"
    exit 1
}

$after = Get-PSDrive -Name InUseTestDrive -ErrorAction SilentlyContinue
if ($null -ne $after) {
    Write-Host "FAIL: in-use drive was not removed with -Force"
    exit 1
}

Write-Host "PASS"
exit 0
