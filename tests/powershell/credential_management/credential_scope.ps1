# vybe-test: powershell/credential_management/credential_scope
$secure = ConvertTo-SecureString 'pass' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('u',$secure)
if (-not $cred) {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
