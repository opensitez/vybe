# vybe-test: powershell/engine_psdefaultparameter_values/scriptblock_accesses_outer_scope_variable
$script:deploymentRegion = "us-west-2"

$global:PSDefaultParameterValues = @{
    "Deploy-Component:Region" = { $script:deploymentRegion }
}

function Deploy-Component {
    [CmdletBinding()]
    param([string]$Region)
    return $Region
}

$firstRun = Deploy-Component

# Mutating outer scope variable must be reflected in next dynamic invocation
$script:deploymentRegion = "eu-central-1"
$secondRun = Deploy-Component

$global:PSDefaultParameterValues.Clear()

if ($firstRun -ne "us-west-2") {
    Write-Host "FAIL: first run expected 'us-west-2', got '$firstRun'"
    exit 1
}

if ($secondRun -ne "eu-central-1") {
    Write-Host "FAIL: second run expected 'eu-central-1' after outer variable mutation, got '$secondRun'"
    exit 1
}

Write-Host "PASS"
exit 0
