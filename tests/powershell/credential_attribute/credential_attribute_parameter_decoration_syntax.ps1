# vybe-test: powershell/credential_attribute/credential_attribute_parameter_decoration_syntax
# Decorating a PSCredential parameter with [System.Management.Automation.CredentialAttribute()] compiles and binds
function ConnectServiceDirect {
    [CmdletBinding()]
    param(
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$AuthCredential
    )
    process {
        return "Authenticated:$($AuthCredential.UserName)"
    }
}

$sec = ConvertTo-SecureString "SecretPass" -AsPlainText -Force
$cred = [System.Management.Automation.PSCredential]::new("svc_worker", $sec)

$result = ConnectServiceDirect -AuthCredential $cred

if ($result -ne "Authenticated:svc_worker") {
    Write-Host "FAIL: direct credential binding failed: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
