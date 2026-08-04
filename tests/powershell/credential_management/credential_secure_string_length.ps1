# vybe-test: powershell/credential_management/credential_secure_string_length
$secure = ConvertTo-SecureString 'pass' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('u',$secure)
if ($cred.Password.Length -ne 4) {
    Write-Host "FAIL: expected 4"
    exit 1
}
Write-Host 'PASS'
exit 0
