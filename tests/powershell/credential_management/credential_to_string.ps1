# vybe-test: powershell/credential_management/credential_to_string
$sec = ConvertTo-SecureString "Pass123" -AsPlainText -Force
$cred = [System.Management.Automation.PSCredential]::new("admin", $sec)
if ($cred.UserName -ne "admin") {
    Write-Host "FAIL: PSCredential username check failed"
    exit 1
}
Write-Host "PASS"
exit 0
