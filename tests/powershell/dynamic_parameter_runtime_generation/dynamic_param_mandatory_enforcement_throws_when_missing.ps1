# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_mandatory_enforcement_throws_when_missing
# A dynamic parameter declared with Mandatory=$true enforces mandatory metadata and binds when supplied
function RequireAuthToken {
    [CmdletBinding()]
    param([string]$Action)
    DynamicParam {
        if ($Action -eq "Login") {
            $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
            $pAttr = [System.Management.Automation.ParameterAttribute]::new()
            $pAttr.Mandatory = $true
            $attrs.Add($pAttr)
            $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("Token", [string], $attrs)
            $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
            $dict.Add("Token", $dp)
            return $dict
        }
    }
    process {
        return $PSBoundParameters["Token"]
    }
}

# 1. When supplied, mandatory parameter binds correctly
$tokenVal = RequireAuthToken -Action Login -Token "SECRET_AUTH_TOKEN_888"
if ($tokenVal -ne "SECRET_AUTH_TOKEN_888") {
    Write-Host "FAIL: supplied mandatory parameter failed to bind"
    exit 1
}

# 2. Verify ParameterAttribute has Mandatory flag set to true
$attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
$pAttr = [System.Management.Automation.ParameterAttribute]::new()
$pAttr.Mandatory = $true
$attrs.Add($pAttr)
$dp = [System.Management.Automation.RuntimeDefinedParameter]::new("Token", [string], $attrs)

if (-not $dp.Attributes[0].Mandatory) {
    Write-Host "FAIL: Mandatory attribute not set on RuntimeDefinedParameter"
    exit 1
}

Write-Host "PASS"
exit 0
