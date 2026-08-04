# vybe-test: powershell/functions/switch_parameter
function Test-Switch {
    param(
        [switch]$Enable
    )
    if ($Enable) {
        return "enabled"
    }
    return "disabled"
}
$result = Test-Switch -Enable
if ($result -ne "enabled") {
    Write-Host "FAIL: expected 'enabled', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
