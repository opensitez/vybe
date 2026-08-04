# vybe-test: powershell/credential_management/credential_password_type
$secure = ConvertTo-SecureString 'pass' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('u',$secure)
if (-not ($cred.Password -is [System.Security.SecureString])) {
    Write-Host "FAIL: expected SecureString password"
    exit 1
}
Write-Host 'PASS'
exit 0
