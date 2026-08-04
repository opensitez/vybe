# vybe-test: powershell/credential_management/credential_hashcode
$secure = ConvertTo-SecureString 'pass' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('u',$secure)
if (-not $cred.GetHashCode()) {
    Write-Host "FAIL: expected hashcode"
    exit 1
}
Write-Host 'PASS'
exit 0
