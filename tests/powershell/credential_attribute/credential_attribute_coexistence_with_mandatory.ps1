# vybe-test: powershell/credential_attribute/credential_attribute_coexistence_with_mandatory
# ParameterAttribute.Mandatory and CredentialAttribute coexist without conflict
function RequireCredentialAuth {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$Credential
    )
    process {
        return "MandatoryAuth:$($Credential.UserName)"
    }
}

$sec = ConvertTo-SecureString "Pass1" -AsPlainText -Force
$cred = [System.Management.Automation.PSCredential]::new("mandatory_admin", $sec)

$res = RequireCredentialAuth -Credential $cred

if ($res -ne "MandatoryAuth:mandatory_admin") {
    Write-Host "FAIL: mandatory credential call failed: '$res'"
    exit 1
}

# Verify Mandatory = true via metadata
$param = (Get-Command RequireCredentialAuth).Parameters["Credential"]
$isMandatory = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] } | Select-Object -ExpandProperty Mandatory

if (-not $isMandatory) {
    Write-Host "FAIL: parameter was not marked Mandatory"
    exit 1
}

Write-Host "PASS"
exit 0
