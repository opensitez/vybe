# vybe-test: powershell/engine_psdefaultparameter_values/failed_validation_emits_warning_and_leaves_unbound
# When a default value violates a validation attribute (e.g. ValidateSet),
# PowerShell writes a warning and leaves the parameter unbound ($null) without terminating.
$global:PSDefaultParameterValues = @{
    "Set-TrafficLight:Color" = "Purple"
}

function Set-TrafficLight {
    [CmdletBinding()]
    param(
        [ValidateSet("Red", "Yellow", "Green")]
        $Color
    )
    return $Color
}

$res = Set-TrafficLight 3>$null

$global:PSDefaultParameterValues.Clear()

# The parameter must have remained unbound ($null) due to failed validation
if ($null -ne $res) {
    Write-Host "FAIL: invalid default value was unexpectedly bound, got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
