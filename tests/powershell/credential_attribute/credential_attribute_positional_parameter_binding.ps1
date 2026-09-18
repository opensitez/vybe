# vybe-test: powershell/credential_attribute/credential_attribute_positional_parameter_binding
# Positional binding works for a PSCredential parameter decorated with CredentialAttribute
function AuthenticatePositional {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$Credential
    )
    process {
        return "PositionalUser:$($Credential.UserName)"
    }
}

$sec = ConvertTo-SecureString "Pass1" -AsPlainText -Force
$cred = [System.Management.Automation.PSCredential]::new("positional_operator", $sec)

$res = AuthenticatePositional $cred

if ($res -ne "PositionalUser:positional_operator") {
    Write-Host "FAIL: positional credential binding failed: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
