# vybe-test: powershell/credential_management/credential_string_conversion
$secure = ConvertTo-SecureString 'pass' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('user',$secure)
if ($cred.Password.Length -ne $secure.Length) {
    Write-Host "FAIL: expected password length match"
    exit 1
}
Write-Host 'PASS'
exit 0
