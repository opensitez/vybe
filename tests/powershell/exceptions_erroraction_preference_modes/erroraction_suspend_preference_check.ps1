# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_suspend_preference_check
$pref = [System.Management.Automation.ActionPreference]::Suspend
if ($pref.ToString() -ne "Suspend") {
    Write-Host "FAIL: Suspend ActionPreference check failed"
    exit 1
}
Write-Host "PASS"
exit 0
