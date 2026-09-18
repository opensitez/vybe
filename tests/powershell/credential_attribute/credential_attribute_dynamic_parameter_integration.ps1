# vybe-test: powershell/credential_attribute/credential_attribute_dynamic_parameter_integration
# CredentialAttribute can be attached to a dynamically generated RuntimeDefinedParameter
function ConnectDynamically {
    [CmdletBinding()]
    param([string]$AuthMode)
    DynamicParam {
        if ($AuthMode -eq "CustomCred") {
            $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
            $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
            $attrs.Add([System.Management.Automation.CredentialAttribute]::new())
            $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("DynamicCred", [pscredential], $attrs)
            $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
            $dict.Add("DynamicCred", $dp)
            return $dict
        }
    }
    process {
        if ($PSBoundParameters.ContainsKey("DynamicCred")) {
            return "DynUser:$($PSBoundParameters['DynamicCred'].UserName)"
        }
        return "NoDynCred"
    }
}

$sec = ConvertTo-SecureString "Pass1" -AsPlainText -Force
$cred = [System.Management.Automation.PSCredential]::new("dyn_user_99", $sec)

$res = ConnectDynamically -AuthMode "CustomCred" -DynamicCred $cred

if ($res -ne "DynUser:dyn_user_99") {
    Write-Host "FAIL: dynamic parameter credential integration failed: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
