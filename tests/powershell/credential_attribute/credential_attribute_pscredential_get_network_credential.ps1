# vybe-test: powershell/credential_attribute/credential_attribute_pscredential_get_network_credential
# PSCredential.GetNetworkCredential().Password recovers the cleartext password from the secure string
function VerifyNetworkPassword {
    [CmdletBinding()]
    param(
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$Cred
    )
    process {
        return $Cred.GetNetworkCredential().Password
    }
}

$sec = ConvertTo-SecureString "PlainTextSecret123" -AsPlainText -Force
$cred = [System.Management.Automation.PSCredential]::new("db_user", $sec)

$plainPassword = VerifyNetworkPassword -Cred $cred

if ($plainPassword -ne "PlainTextSecret123") {
    Write-Host "FAIL: recovered password mismatch: '$plainPassword'"
    exit 1
}

Write-Host "PASS"
exit 0
