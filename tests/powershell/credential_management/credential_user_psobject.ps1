# vybe-test: powershell/credential_management/credential_user_psobject
$secure = ConvertTo-SecureString 'pass' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('u',$secure)
if ($cred.UserName -ne 'u') {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
