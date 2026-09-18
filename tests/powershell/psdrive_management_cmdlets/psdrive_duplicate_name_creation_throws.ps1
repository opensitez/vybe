# vybe-test: powershell/psdrive_management_cmdlets/psdrive_duplicate_name_creation_throws
# Calling New-PSDrive with a drive name that already exists throws an exception
$firstDrive = New-PSDrive -Name UniqueDriveName99 -PSProvider FileSystem -Root $pwd.Path

$threwError = $false
try {
    New-PSDrive -Name UniqueDriveName99 -PSProvider FileSystem -Root $pwd.Path -ErrorAction Stop
} catch {
    $threwError = $true
}

Remove-PSDrive -Name UniqueDriveName99

if (-not $threwError) {
    Write-Host "FAIL: creating drive with duplicate name did not throw an error"
    exit 1
}

Write-Host "PASS"
exit 0
