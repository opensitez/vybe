# vybe-test: powershell/credential_attribute/credential_attribute_pscredential_username_property
# PSCredential.UserName property is accurately preserved through parameter passing
function GetBoundUserName {
    [CmdletBinding()]
    param(
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$Cred
    )
    process {
        return $Cred.UserName
    }
}

$sec = ConvertTo-SecureString "Pass123!" -AsPlainText -Force
$cred = [System.Management.Automation.PSCredential]::new("corp_domain\service_svc", $sec)

$userName = GetBoundUserName -Cred $cred

if ($userName -ne "corp_domain\service_svc") {
    Write-Host "FAIL: UserName mismatch, expected 'corp_domain\service_svc', got '$userName'"
    exit 1
}

Write-Host "PASS"
exit 0
