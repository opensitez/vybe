# vybe-test: powershell/credential_attribute/credential_attribute_get_command_metadata_inspection
# Get-Command reflects the presence of CredentialAttribute on the parameter's Attributes collection
function InspectParamMetadata {
    [CmdletBinding()]
    param(
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$AccountCredentials
    )
    process { "ok" }
}

$cmd = Get-Command InspectParamMetadata
$param = $cmd.Parameters["AccountCredentials"]
$hasCredAttr = ($param.Attributes | Where-Object { $_ -is [System.Management.Automation.CredentialAttribute] }) -ne $null

if (-not $hasCredAttr) {
    Write-Host "FAIL: CredentialAttribute not found in parameter Attributes collection"
    exit 1
}

Write-Host "PASS"
exit 0
