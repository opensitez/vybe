# vybe-test: powershell/engine_psdefaultparameter_values/wildcard_parameter_matches_parameter_name
$global:PSDefaultParameterValues = @{ "Show-TestBanner:M*" = "WelcomeBanner" }

function Show-TestBanner {
    [CmdletBinding()]
    param([string]$Message)
    return $Message
}

$res = Show-TestBanner

$global:PSDefaultParameterValues.Clear()

if ($res -ne "WelcomeBanner") {
    Write-Host "FAIL: wildcard parameter pattern M* failed to bind to `$Message, got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
