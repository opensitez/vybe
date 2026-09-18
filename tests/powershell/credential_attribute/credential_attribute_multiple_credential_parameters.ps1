# vybe-test: powershell/credential_attribute/credential_attribute_multiple_credential_parameters
# An advanced function can declare multiple independent CredentialAttribute decorated parameters
function CopyResourceWithCredentials {
    [CmdletBinding()]
    param(
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$SourceCredential,

        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$TargetCredential
    )
    process {
        return "Source:$($SourceCredential.UserName)->Target:$($TargetCredential.UserName)"
    }
}

$sec = ConvertTo-SecureString "Pass1" -AsPlainText -Force
$src = [System.Management.Automation.PSCredential]::new("source_operator", $sec)
$tgt = [System.Management.Automation.PSCredential]::new("target_operator", $sec)

$res = CopyResourceWithCredentials -SourceCredential $src -TargetCredential $tgt

if ($res -ne "Source:source_operator->Target:target_operator") {
    Write-Host "FAIL: multiple credential parameter binding failed: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
