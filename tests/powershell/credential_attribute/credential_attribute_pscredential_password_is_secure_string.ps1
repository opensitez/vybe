# vybe-test: powershell/credential_attribute/credential_attribute_pscredential_password_is_secure_string
# PSCredential.Password property is typed as System.Security.SecureString
function VerifySecurePassword {
    [CmdletBinding()]
    param(
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$Cred
    )
    process {
        return $Cred.Password.GetType().FullName
    }
}

$sec = ConvertTo-SecureString "SecureValue99" -AsPlainText -Force
$cred = [System.Management.Automation.PSCredential]::new("auditor", $sec)

$passwordType = VerifySecurePassword -Cred $cred

if ($passwordType -ne "System.Security.SecureString") {
    Write-Host "FAIL: expected SecureString type, got '$passwordType'"
    exit 1
}

Write-Host "PASS"
exit 0
