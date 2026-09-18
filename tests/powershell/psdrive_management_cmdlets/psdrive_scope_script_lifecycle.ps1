# vybe-test: powershell/psdrive_management_cmdlets/psdrive_scope_script_lifecycle
# -Scope Script mounts the drive visible across the entire script execution scope
$driveName = "ScriptScopedDrive"

& {
    $drive = New-PSDrive -Name $driveName -PSProvider FileSystem -Root $pwd.Path -Scope Script
    if ($null -eq $drive) {
        Write-Host "FAIL: drive creation in child block failed"
        exit 1
    }
}

$visibleInScript = Get-PSDrive -Name $driveName -ErrorAction SilentlyContinue
Remove-PSDrive -Name $driveName

if ($null -eq $visibleInScript) {
    Write-Host "FAIL: script-scoped drive was not visible outside child block"
    exit 1
}

Write-Host "PASS"
exit 0
