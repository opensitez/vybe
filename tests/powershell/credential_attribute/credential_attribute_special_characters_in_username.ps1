# vybe-test: powershell/credential_attribute/credential_attribute_special_characters_in_username
# Usernames with backslashes (domain format) and @ signs (UPN format) are preserved accurately
function CheckDomainUsername {
    [CmdletBinding()]
    param(
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$Credential
    )
    process {
        return "User:$($Credential.UserName)"
    }
}

$sec = ConvertTo-SecureString "Pass1" -AsPlainText -Force
$domainCred = [System.Management.Automation.PSCredential]::new("CORP_ENT\admin.root", $sec)
$upnCred = [System.Management.Automation.PSCredential]::new("service_account@corp.domain.internal", $sec)

$res1 = CheckDomainUsername -Credential $domainCred
$res2 = CheckDomainUsername -Credential $upnCred

if ($res1 -ne "User:CORP_ENT\admin.root") {
    Write-Host "FAIL: domain backslash format failed: '$res1'"
    exit 1
}

if ($res2 -ne "User:service_account@corp.domain.internal") {
    Write-Host "FAIL: UPN format failed: '$res2'"
    exit 1
}

Write-Host "PASS"
exit 0
