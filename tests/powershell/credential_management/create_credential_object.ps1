# vybe-test: powershell/credential_management/create_credential_object
$secure = ConvertTo-SecureString 'pass' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('user',$secure)
if (-not $cred) {
    Write-Host "FAIL: expected PSCredential"
    exit 1
}
Write-Host 'PASS'
exit 0
