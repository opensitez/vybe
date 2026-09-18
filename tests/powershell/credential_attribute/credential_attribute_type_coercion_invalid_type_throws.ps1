# vybe-test: powershell/credential_attribute/credential_attribute_type_coercion_invalid_type_throws
# Passing an incompatible type such as an integer or hashtable throws a parameter binding exception
function EnforceCredentialType {
    [CmdletBinding()]
    param(
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$Credential
    )
    process {
        return $Credential.UserName
    }
}

$threw = $false
try {
    # Passing an integer 12345 cannot be transformed into PSCredential
    EnforceCredentialType -Credential 12345 -ErrorAction Stop
} catch [System.Management.Automation.ParameterBindingException] {
    $threw = $true
}

if (-not $threw) {
    Write-Host "FAIL: incompatible integer input did not throw ParameterBindingException"
    exit 1
}

Write-Host "PASS"
exit 0
