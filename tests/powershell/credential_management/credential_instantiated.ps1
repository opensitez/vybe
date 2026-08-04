# vybe-test: powershell/credential_management/credential_instantiated
$secure = ConvertTo-SecureString 'pass' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('u',$secure)
if (-not $cred.GetType().Name) {
    Write-Host "FAIL: expected credential type"
    exit 1
}
Write-Host 'PASS'
exit 0
