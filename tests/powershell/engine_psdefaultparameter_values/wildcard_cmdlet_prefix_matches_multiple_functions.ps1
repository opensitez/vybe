# vybe-test: powershell/engine_psdefaultparameter_values/wildcard_cmdlet_prefix_matches_multiple_functions
$global:PSDefaultParameterValues = @{ "Audit-*:Environment" = "Staging" }

function Audit-Database {
    [CmdletBinding()]
    param([string]$Environment)
    return $Environment
}

function Audit-Network {
    [CmdletBinding()]
    param([string]$Environment)
    return $Environment
}

$dbEnv = Audit-Database
$netEnv = Audit-Network

$global:PSDefaultParameterValues.Clear()

if ($dbEnv -ne "Staging" -or $netEnv -ne "Staging") {
    Write-Host "FAIL: wildcard cmdlet pattern failed to bind to both functions, got db='$dbEnv', net='$netEnv'"
    exit 1
}

Write-Host "PASS"
exit 0
