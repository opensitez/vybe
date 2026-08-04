# vybe-test: powershell/credential_management/credential_to_string
$secure = ConvertTo-SecureString 'pass' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('u',$secure)
if ($cred.ToString() -notlike '*System.Security*') {
    Write-Host "FAIL: expected PSCredential string"
    exit 1
}
Write-Host 'PASS'
exit 0
