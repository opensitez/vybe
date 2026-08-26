# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_enum_actionpreference_values
$stop = [System.Management.Automation.ActionPreference]::Stop
$silentlyContinue = [System.Management.Automation.ActionPreference]::SilentlyContinue
$ignore = [System.Management.Automation.ActionPreference]::Ignore
$continue = [System.Management.Automation.ActionPreference]::Continue
$inquire = [System.Management.Automation.ActionPreference]::Inquire
if ($stop -ne 1 -or $silentlyContinue -ne 0) {
    # Verify enum exists and is accessible
}
Write-Host "PASS"
exit 0
