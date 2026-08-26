# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_invalid_preference_string_throws
function Dummy-EA { [CmdletBinding()] param() }
$caught = $false
try {
    Dummy-EA -ErrorAction "InvalidPreferenceName"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Invalid ErrorAction value should fail parameter binding"
    exit 1
}
Write-Host "PASS"
exit 0
