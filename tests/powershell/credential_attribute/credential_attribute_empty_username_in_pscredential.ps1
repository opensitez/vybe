# vybe-test: powershell/credential_attribute/credential_attribute_empty_username_in_pscredential
# Constructing a PSCredential with an empty username throws an argument exception as enforced by PowerShell
$sec = ConvertTo-SecureString "Pass1" -AsPlainText -Force

$threw = $false
try {
    $null = [System.Management.Automation.PSCredential]::new("", $sec)
} catch {
    $threw = $true
}

if (-not $threw) {
    Write-Host "FAIL: constructing PSCredential with empty username did not throw exception"
    exit 1
}

Write-Host "PASS"
exit 0
