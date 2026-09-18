# vybe-test: powershell/credential_attribute/credential_attribute_value_from_pipeline_by_property_name
# Pipeline objects with a Credential property bind correctly via ValueFromPipelineByPropertyName
function BindCredByPropertyName {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$Credential
    )
    process {
        return "PropBound:$($Credential.UserName)"
    }
}

$sec = ConvertTo-SecureString "Pass1" -AsPlainText -Force
$cred = [System.Management.Automation.PSCredential]::new("pipeline_prop_user", $sec)
$containerObj = [PSCustomObject]@{
    Credential = $cred
    ClusterName = "ClusterBeta"
}

$res = $containerObj | BindCredByPropertyName

if ($res -ne "PropBound:pipeline_prop_user") {
    Write-Host "FAIL: ValueFromPipelineByPropertyName binding mismatch: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
