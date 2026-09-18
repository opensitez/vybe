# vybe-test: powershell/credential_attribute/credential_attribute_null_credential_handling
# Passing $null to an optional CredentialAttribute parameter binds null cleanly without throwing
function AllowNullCredential {
    [CmdletBinding()]
    param(
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$Credential = $null
    )
    process {
        $isNull = $Credential -eq $null
        return "IsNull:$isNull"
    }
}

$resDefault = AllowNullCredential
$resExplicitNull = AllowNullCredential -Credential $null

if ($resDefault -ne "IsNull:True") {
    Write-Host "FAIL: default null credential failed: '$resDefault'"
    exit 1
}

if ($resExplicitNull -ne "IsNull:True") {
    Write-Host "FAIL: explicit null credential failed: '$resExplicitNull'"
    exit 1
}

Write-Host "PASS"
exit 0
