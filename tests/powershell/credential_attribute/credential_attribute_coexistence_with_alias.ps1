# vybe-test: powershell/credential_attribute/credential_attribute_coexistence_with_alias
# AliasAttribute coexists with CredentialAttribute allowing parameter resolution via alias
function AuthenticateViaAlias {
    [CmdletBinding()]
    param(
        [System.Management.Automation.CredentialAttribute()]
        [Alias("Auth", "UserCred")]
        [pscredential]$Credential
    )
    process {
        return "AliasUser:$($Credential.UserName)"
    }
}

$sec = ConvertTo-SecureString "Pass1" -AsPlainText -Force
$cred = [System.Management.Automation.PSCredential]::new("alias_user", $sec)

$res1 = AuthenticateViaAlias -Auth $cred
$res2 = AuthenticateViaAlias -UserCred $cred

if ($res1 -ne "AliasUser:alias_user" -or $res2 -ne "AliasUser:alias_user") {
    Write-Host "FAIL: alias binding failed: '$res1', '$res2'"
    exit 1
}

Write-Host "PASS"
exit 0
