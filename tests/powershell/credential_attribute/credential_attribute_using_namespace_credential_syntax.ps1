# vybe-test: powershell/credential_attribute/credential_attribute_using_namespace_credential_syntax
# 'using namespace System.Management.Automation' enables short [Credential()] parameter decoration syntax
using namespace System.Management.Automation

function ConnectServiceWithNamespace {
    [CmdletBinding()]
    param(
        [Credential()]
        [pscredential]$Credential
    )
    process {
        return "ConnectedUser:$($Credential.UserName)"
    }
}

$sec = ConvertTo-SecureString "Pass#456" -AsPlainText -Force
$cred = [PSCredential]::new("cluster_admin", $sec)

$res = ConnectServiceWithNamespace -Credential $cred

if ($res -ne "ConnectedUser:cluster_admin") {
    Write-Host "FAIL: short [Credential()] attribute syntax failed: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
