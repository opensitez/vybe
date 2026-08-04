# vybe-test: powershell/credential_management/credential_user_name
$secure = ConvertTo-SecureString 'pass' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('user',$secure)
if ($cred.UserName -ne 'user') {
    Write-Host "FAIL: expected user"
    exit 1
}
Write-Host 'PASS'
exit 0
