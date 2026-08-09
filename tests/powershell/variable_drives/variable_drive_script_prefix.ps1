# vybe-test: powershell/variable_drives/variable_drive_script_prefix
$script:scriptLevel = "ScriptScoped"
function Check-ScriptDrive {
    return $script:scriptLevel
}
$res = Check-ScriptDrive
if ($res -ne "ScriptScoped") {
    Write-Host "FAIL: \$script: drive prefix in function scope expected ScriptScoped, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
